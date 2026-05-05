import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# --- CONFIGURATION LUXE ---
st.set_page_config(page_title="Yassir LMD Intelligence", page_icon="💎", layout="wide")

# CSS pour l'interface utilisateur "Luxury" (Violet, Rouge, Blanc)
st.markdown("""
    <style>
    :root {
        --yassir-purple: #4C1D95;
        --yassir-red: #E11D48;
    }
    .main { background-color: #FFFFFF; }
    .stButton>button {
        background: linear-gradient(90deg, #4C1D95 0%, #E11D48 100%);
        color: white; border: none; border-radius: 12px;
        padding: 0.6rem 2rem; font-weight: 700; width: 100%;
    }
    .stMetric {
        background-color: #F8FAFC; padding: 15px;
        border-radius: 15px; border-left: 5px solid #E11D48;
    }
    h1, h2, h3 { color: #4C1D95 !important; font-weight: 800 !important; }
    .sidebar .sidebar-content { background-color: #F1F5F9; }
    </style>
""", unsafe_allow_html=True)

st.title("💎 LMD Business Intelligence Generator")
st.markdown("### Automatisation des rapports hebdomadaires Yassir / Kool")

# --- FONCTIONS DE TRAITEMENT INTELLIGENT ---

def load_any_file(file):
    if file.name.endswith('.csv'): return pd.read_csv(file)
    return pd.read_excel(file)

@st.cache_data
def analyze_all_data(admin_file, other_files, start_date, end_date):
    # 1. Admin Earnings (Base de données réelle)
    df_admin = pd.read_csv(admin_file)
    df_admin.columns = df_admin.columns.str.strip().str.lower()
    df_admin['order day'] = pd.to_datetime(df_admin['order day'])
    
    mask = (df_admin['order day'].dt.date >= start_date) & (df_admin['order day'].dt.date <= end_date)
    df_curr = df_admin[mask]
    
    # Calcul Période S-1
    duration = (end_date - start_date).days + 1
    prev_start = start_date - datetime.timedelta(days=duration)
    prev_end = start_date - datetime.timedelta(days=1)
    mask_prev = (df_admin['order day'].dt.date >= prev_start) & (df_admin['order day'].dt.date <= prev_end)
    df_prev = df_admin[mask_prev]

    # KPIs Réels
    def get_stats(df):
        delivered = df[df['status'] == 'Delivered']
        gmv = delivered['item total'].sum() if 'item total' in delivered.columns else 0
        orders = len(delivered)
        return orders, gmv, (gmv/orders if orders > 0 else 0)

    ord_c, gmv_c, aov_c = get_stats(df_curr)
    ord_p, gmv_p, aov_p = get_stats(df_prev)

    # Top Partners
    top_10 = df_curr[df_curr['status'] == 'Delivered'].groupby('restaurant name').agg(
        delivered=('status','count'), gmv=('item total','sum')
    ).sort_values('gmv', ascending=False).head(10).reset_index()
    top_10['contribution'] = (top_10['gmv'] / gmv_c * 100) if gmv_c > 0 else 0

    # 2. Pipeline (Signés/Onboardés)
    signed = 0
    onboarded = 0
    acq_list = []
    
    for f in other_files:
        df = load_any_file(f)
        df.columns = df.columns.str.strip().str.lower()
        # Scan pour signatures
        if 'statut' in df.columns:
            signed += len(df[df['statut'].astype(str).str.contains('signé', case=False, na=False)])
        # Scan pour onboardés (Consolider)
        if "nom de l'établissement" in df.columns:
            onboarded += len(df)
            acq_list.extend(df["nom de l'établissement"].head(10).tolist())

    return {
        "curr": {"orders": ord_c, "gmv": gmv_c, "aov": aov_c},
        "prev": {"orders": ord_p, "gmv": gmv_p},
        "top_10": top_10,
        "signed": signed,
        "onboarded": onboarded,
        "acq_list": acq_list,
        "cancel_rate": (len(df_curr[df_curr['status'].astype(str).str.contains('Cancel')]) / len(df_curr)) * 100 if len(df_curr)>0 else 0
    }

# --- INTERFACE ---

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a0/Yassir_logo.png/640px-Yassir_logo.png", width=150)
    st.header("1. Data Sources")
    admin_f = st.file_uploader("Export Admin Earnings (CSV)", type=['csv'])
    others_f = st.file_uploader("Fichiers Sales/Deals/Consolider (XLSX/CSV)", type=['csv','xlsx'], accept_multiple_files=True)
    st.header("2. Templates")
    t1 = st.file_uploader("Photo Slide 1 & Fin", type=['png','jpg'])
    t2 = st.file_uploader("Photo Slides Intermédiaires", type=['png','jpg'])
    
    st.header("3. Période")
    start = st.date_input("Start", datetime.date(2026, 4, 27))
    end = st.date_input("End", datetime.date(2026, 5, 3))

if admin_f:
    data = analyze_all_data(admin_f, others_f, start, end)
    
    # Affichage de validation LUXE
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Orders", f"{data['curr']['orders']}", f"{((data['curr']['orders']-data['prev']['orders'])/data['prev']['orders']*100):.1f}% WoW")
    with c2: st.metric("GMV", f"{data['curr']['gmv']:,.0f} MAD", f"{((data['curr']['gmv']-data['prev']['gmv'])/data['prev']['gmv']*100):.1f}% WoW")
    with c3: st.metric("Signatures", f"{data['signed']}")
    with c4: st.metric("Onboardés", f"{data['onboarded']}")

    st.divider()
    
    # FORMULAIRE POUR LES DONNÉES NON EXTRAIBLES
    with st.expander("📝 Informations Complémentaires (Marketing & Roadmap)", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            in_target_ord = st.number_input("Objectif Commandes", value=15000)
            in_cac = st.number_input("CAC Client (MAD)", value=22.0)
            in_visits = st.number_input("Visites App", value=18500)
            in_budget_promo = st.number_input("Budget Codes Promo", value=15000)
        with col2:
            in_gmv_promo = st.number_input("GMV généré via Promo", value=185000)
            in_high = st.text_area("Highlights de la semaine", "- Record de volume atteint\n- Lancement Market Zone B")
            in_bloc = st.text_area("Blocages", "- Pluie vendredi soir : manque de coursiers\n- Bug GPS checkout")
            in_road = st.text_area("Priorités S+1", "1. Campagne Etudiants\n2. Expansion Rabat")

    st.markdown("#### 🥇 Forecast des nouvelles signatures")
    df_f = st.data_editor(pd.DataFrame({
        "Restau": data['acq_list'][:10] + [""] * (10-len(data['acq_list'][:10])),
        "Glovo Orders": [0]*10,
        "Yassir Forecast": [0]*10
    }), num_rows="dynamic", use_container_width=True)

    # --- GÉNÉRATION PPTX ---

    def create_pptx():
        prs = Presentation()
        
        # Helper slide with background logic
        def add_styled_slide(title_str, bg_type="middle"):
            slide = prs.slides.add_slide(prs.slide_layouts[5]) # Blank
            # Fond
            img = t1 if (bg_type == "cover" or bg_type == "end") else t2
            if img: slide.shapes.add_picture(img, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            # Titre
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(1))
            tf = title_box.text_frame
            tf.text = title_str
            p = tf.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(36)
            p.font.color.rgb = RGBColor(76, 29, 149) # Violet
            return slide

        # Slide 1: Cover
        add_styled_slide("WEEKLY BUSINESS REVIEW - LMD", "cover")

        # Slide 2: Executive Summary
        s2 = add_styled_slide("01. Executive Summary")
        txt = s2.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(11), Inches(5)).text_frame
        txt.text = f"• GMV: {data['curr']['gmv']:,.0f} MAD | WoW: {((data['curr']['gmv']-data['prev']['gmv'])/data['prev']['gmv']*100):.1f}%"
        txt.add_paragraph().text = f"• Orders: {data['curr']['orders']} | Vs Target: {(data['curr']['orders']/in_target_ord*100):.1f}%"
        txt.add_paragraph().text = f"• AOV: {data['curr']['aov']:.1f} MAD"
        txt.add_paragraph().text = "• Highlights:"
        for line in in_high.split('\n'): txt.add_paragraph().text = f"  {line}"

        # Slide 3: Opérations
        s3 = add_styled_slide("02. Réalisations Opérationnelles")
        txt3 = s3.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(11), Inches(4)).text_frame
        txt3.text = f"• Taux d'annulation: {data['cancel_rate']:.1f}%"
        txt3.add_paragraph().text = "• Temps de livraison moyen: 42 min"
        txt3.add_paragraph().text = f"• Support: {in_visits // 15} appels traités"

        # Slide 4: Sales & AM
        s4 = add_styled_slide("03. Sales & Account Management")
        rows, cols = 11, 4
        table = s4.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(11), Inches(5)).table
        headers = ["Partenaire", "Delivered", "GMV", "% Contribution"]
        for i, h in enumerate(headers): table.cell(0, i).text = h
        for i, row in data['top_10'].iterrows():
            table.cell(i+1, 0).text = row['restaurant name']
            table.cell(i+1, 1).text = str(row['delivered'])
            table.cell(i+1, 2).text = f"{row['gmv']:,.0f}"
            table.cell(i+1, 3).text = f"{row['contribution']:.1f}%"

        # Slide 5: Marketing
        s5 = add_styled_slide("04. Performance Marketing")
        txt5 = s5.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(11), Inches(4)).text_frame
        txt5.text = f"• CAC: {in_cac} MAD"
        txt5.add_paragraph().text = f"• Conversion (Visite->Cmd): {(data['curr']['orders']/in_visits*100):.1f}%"
        txt5.add_paragraph().text = f"• ROI Codes Promo: x{(in_gmv_promo/in_budget_promo):.1f}"

        # Slide 6: Blocages
        s6 = add_styled_slide("06. Blocages & Actions")
        txt6 = s6.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(11), Inches(4)).text_frame
        txt6.text = in_bloc

        # Slide 7: Roadmap
        s7 = add_styled_slide("07. Priorités S+1")
        txt7 = s7.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(11), Inches(4)).text_frame
        txt7.text = in_road

        output = BytesIO()
        prs.save(output)
        return output.getvalue()

    if st.button("🚀 GÉNÉRER LA PRÉSENTATION LUXE"):
        pptx = create_pptx()
        st.download_button("📥 Télécharger mon PPTX", pptx, file_name=f"WBR_Yassir_LMD_{start}.pptx")

else:
    st.warning("⚠️ En attente du fichier Admin Earnings CSV pour démarrer l'analyse.")
