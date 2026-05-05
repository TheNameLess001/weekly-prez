import streamlit as st
import pandas as pd
import datetime
import base64

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Générateur WBR LMD", page_icon="🚀", layout="wide")

st.title("🚀 Générateur Weekly Business Review (Yassir / Kool)")
st.markdown("Importez vos exports, validez les chiffres, saisissez les prévisions et générez la présentation.")

# --- FONCTIONS UTILITAIRES ---
@st.cache_data
def process_admin_file(file):
    df = pd.read_csv(file)
    df.columns = df.columns.str.strip().str.lower()
    if 'order day' in df.columns:
        df['order day'] = pd.to_datetime(df['order day'], errors='coerce')
    return df

@st.cache_data
def process_aux_files(files):
    df_sales_list = []
    df_cons_list = []
    
    for f in files:
        try:
            df = pd.read_csv(f)
            cols = df.columns.str.strip().str.lower()
            df.columns = cols
            
            # Reconnaissance du type de fichier
            if 'statut' in cols and 'segment ' in cols or 'segment' in cols:
                df_sales_list.append(df)
            elif "nom de l'établissement" in cols or "nom de l'etablissement" in cols:
                df_cons_list.append(df)
        except Exception:
            continue
            
    df_sales = pd.concat(df_sales_list, ignore_index=True) if df_sales_list else pd.DataFrame()
    df_cons = pd.concat(df_cons_list, ignore_index=True) if df_cons_list else pd.DataFrame()
    
    # Nettoyage date Consolider
    if not df_cons.empty and 'date' in df_cons.columns:
        df_cons['date'] = pd.to_datetime(df_cons['date'], errors='coerce')
        
    return df_sales, df_cons

def calculate_kpis(df):
    if df is None or df.empty:
        return {'orders': 0, 'gmv': 0, 'aov': 0, 'cancel_rate': 0, 'avg_time': 0, 'top_partners': [], 'cancel_reasons': [], 'total_valid': 0}
    
    total_valid = len(df)
    df_delivered = df[df['status'] == 'Delivered']
    delivered_count = len(df_delivered)
    
    col_gmv = 'item total' if 'item total' in df.columns else 'payable amount'
    gmv = df_delivered[col_gmv].sum() if col_gmv in df.columns else 0
    
    aov = gmv / delivered_count if delivered_count > 0 else 0
    
    col_time = 'delivery time(m)'
    avg_time = df_delivered[col_time].mean() if col_time in df_delivered.columns else 0
        
    df_cancelled = df[df['status'].astype(str).str.contains('Cancel', na=False, case=False)]
    cancelled_count = len(df_cancelled)
    cancel_rate = (cancelled_count / total_valid) * 100 if total_valid > 0 else 0
    
    cancel_reasons = []
    if 'cancellation reason' in df.columns and cancelled_count > 0:
        reasons = df_cancelled['cancellation reason'].value_counts()
        cancel_reasons = [(r, c) for r, c in reasons.items()][:3]

    top_partners = []
    if 'restaurant name' in df.columns:
        top_partners = df_delivered.groupby('restaurant name').agg(
            orders=('status', 'count'), gmv=(col_gmv, 'sum')
        ).reset_index().sort_values(by='gmv', ascending=False).head(10).to_dict('records')

    return {'orders': delivered_count, 'gmv': gmv, 'aov': aov, 'avg_time': avg_time, 'cancel_rate': cancel_rate, 'top_partners': top_partners, 'cancel_reasons': cancel_reasons, 'total_valid': total_valid}

def get_trend(curr, prev, is_pct=False, reverse_logic=False):
    if prev == 0 or pd.isna(prev): return "N/A"
    diff = curr - prev if is_pct else ((curr - prev) / prev) * 100
    
    is_good = diff > 0
    if reverse_logic: is_good = not is_good
    
    color = "#10B981" if is_good else "#EF4444"
    sign = "+" if diff > 0 else ""
    return f"<span style='color:{color}; font-weight:bold; font-size:14px;'>{sign}{diff:.1f}%</span>"

def encode_image(upload):
    if upload is not None:
        return f"data:{upload.type};base64,{base64.b64encode(upload.read()).decode()}"
    return ""

# --- INTERFACE UTILISATEUR ---
st.sidebar.header("📁 1. Fichiers Sources")
admin_file = st.sidebar.file_uploader("Upload CSV (Admin Earnings) *Requis", type=['csv'])
aux_files = st.sidebar.file_uploader("Upload autres exports (Sales Sprint, Consolider, Deals...)", type=['csv'], accept_multiple_files=True)
template_file = st.sidebar.file_uploader("Template Slide de fond (PNG/JPG)", type=['png', 'jpg', 'jpeg'])

if admin_file:
    df_admin = process_admin_file(admin_file)
    df_sales, df_cons = process_aux_files(aux_files) if aux_files else (pd.DataFrame(), pd.DataFrame())
    
    st.sidebar.header("📅 2. Période d'analyse")
    min_d = df_admin['order day'].min().date() if not pd.isna(df_admin['order day'].min()) else datetime.date.today()
    max_d = df_admin['order day'].max().date() if not pd.isna(df_admin['order day'].max()) else datetime.date.today()
    
    start_date = st.sidebar.date_input("Date de début", min_d)
    end_date = st.sidebar.date_input("Date de fin", max_d)
    
    # WoW Calc
    duration = (end_date - start_date).days + 1
    prev_end = start_date - datetime.timedelta(days=1)
    prev_start = prev_end - datetime.timedelta(days=duration - 1)
    
    mask_curr = (df_admin['order day'].dt.date >= start_date) & (df_admin['order day'].dt.date <= end_date)
    mask_prev = (df_admin['order day'].dt.date >= prev_start) & (df_admin['order day'].dt.date <= prev_end)
    
    kpis = calculate_kpis(df_admin[mask_curr])
    kpis_prev = calculate_kpis(df_admin[mask_prev])
    
    # Calcul Signés / Onboardés Automatique
    signed_count = 0
    onboarded_count = 0
    recent_restaurants = []
    
    if not df_sales.empty and 'statut' in df_sales.columns:
        # Approximation : on compte les "Signé" dans les fichiers
        signed_count = len(df_sales[df_sales['statut'].astype(str).str.contains('Signé', case=False, na=False)])
        
    if not df_cons.empty and 'date' in df_cons.columns:
        df_cons_curr = df_cons[(df_cons['date'].dt.date >= start_date) & (df_cons['date'].dt.date <= end_date)]
        onboarded_count = len(df_cons_curr)
        if "nom de l'établissement" in df_cons_curr.columns:
            recent_restaurants = df_cons_curr["nom de l'établissement"].dropna().unique().tolist()[:15]
            
    # --- ONGLETS ---
    tab1, tab2, tab3 = st.tabs(["📊 Validation KPIs", "✍️ Saisie Manuelle & Tableaux", "📤 Génération HTML"])

    with tab1:
        st.subheader(f"Chiffres extraits du {start_date.strftime('%d/%m/%Y')} au {end_date.strftime('%d/%m/%Y')}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Commandes Livrées", f"{kpis['orders']:,}")
        c2.metric("GMV (MAD)", f"{kpis['gmv']:,.0f}")
        c3.metric("Panier Moyen (AOV)", f"{kpis['aov']:,.1f} MAD")
        c4.metric("Taux d'annulation", f"{kpis['cancel_rate']:.1f}%")
        
        st.write("---")
        c_a, c_b = st.columns(2)
        with c_a:
            st.write("**🏆 Top 10 Partenaires**")
            st.dataframe(pd.DataFrame(kpis['top_partners']))
        with c_b:
            st.write(f"**Nouveaux Signés détectés :** {signed_count}")
            st.write(f"**Nouveaux Onboardés détectés :** {onboarded_count}")
            st.write("**Top 3 Annulations :**")
            for r, count in kpis['cancel_reasons']: st.write(f"- {r} ({count} cmds)")

    with tab2:
        st.info("Remplissez les informations manquantes pour finaliser les 7 slides.")
        
        with st.expander("Slide 1 & 2 : Executive Summary & Opérations", expanded=True):
            col1, col2, col3, col4 = st.columns(4)
            in_req = col1.number_input("Total Requests reçues", value=16800)
            in_budget_gmv = col2.number_input("Budget GMV", value=1150000)
            in_freq = col3.number_input("Fréquence", value=2.8, step=0.1)
            in_weekdays = col4.number_input("Weekdays (%)", value=65)
            
            st.write("---")
            col5, col6, col7, col8 = st.columns(4)
            in_t_food = col5.number_input("Tps Moyen Food (m)", value=42)
            in_t_groc = col6.number_input("Tps Moyen Groceries (m)", value=35)
            in_batch = col7.number_input("Batching rate", value=1.3, step=0.1)
            in_auto = col8.number_input("Auto-dispatch (%)", value=94)
            
            c_a, c_b, c_c = st.columns(3)
            in_t_acc = c_a.number_input("Temps Acceptation (m)", value=3)
            in_t_pick = c_b.number_input("Temps Pickup (m)", value=12)
            in_t_traj = c_c.number_input("Temps Trajet (m)", value=27)
            
            c_d, c_e, c_f, c_g = st.columns(4)
            in_fleet = c_d.number_input("Coursiers actifs", value=450)
            in_mix = c_e.number_input("Mix 3PL (%)", value=65)
            in_cpo = c_f.number_input("CPO (MAD)", value=12.5, step=0.1)
            in_calls = c_g.number_input("Appels Support", value=1250)

        with st.expander("Slide 3 : Sales, AM & Forecast Nouveaux", expanded=True):
            st.write("📈 **Pipeline & Account Management**")
            cc1, cc2, cc3, cc4 = st.columns(4)
            in_signed = cc1.number_input("Partenaires Signés", value=signed_count if signed_count > 0 else 146)
            in_onboarded = cc2.number_input("Partenaires Onboardés", value=onboarded_count if onboarded_count > 0 else 110)
            in_churn = cc3.number_input("Taux Churn (%)", value=2.1, step=0.1)
            in_kool = cc4.number_input("Exclusivités KOOL (% GMV)", value=42)
            
            in_topdeal = st.text_input("Top Deal (Highlight exclusif)", value="McDonald's (Exclusivité)")
            
            st.write("---")
            st.write("🥇 **Orders Forecast des nouvelles signatures**")
            st.write("Saisissez les prévisions pour les nouveaux partenariats.")
            
            # Pré-remplir avec les récents
            init_restos = recent_restaurants + [""] * (10 - len(recent_restaurants)) if len(recent_restaurants) < 10 else recent_restaurants[:10]
            
            if 'df_forecast' not in st.session_state:
                st.session_state.df_forecast = pd.DataFrame({
                    "Restaurant": init_restos,
                    "Glovo Orders (Estim.)": [""] * 10,
                    "Yassir Forecast": [""] * 10
                })
            
            edited_forecast = st.data_editor(st.session_state.df_forecast, num_rows="dynamic", use_container_width=True)

        with st.expander("Slide 4 : Growth & Marketing"):
            c1, c2, c3, c4 = st.columns(4)
            in_visits = c1.number_input("Visites App", value=18500)
            in_installs = c2.number_input("Nouveaux Téléchargements", value=4500)
            in_ftu = c3.number_input("Nouveaux Utilisateurs (FTU)", value=2100)
            in_cac = c4.number_input("CAC Client (MAD)", value=22)
            
            c5, c6, c7 = st.columns(3)
            in_campaign = c5.text_input("Top Campagne / Code promo", value="TRIPLE30")
            in_budget_promo = c6.number_input("Budget Alloué (MAD)", value=15000)
            in_gmv_promo = c7.number_input("GMV Généré par Promo", value=185000)

        with st.expander("Slides 5, 6 & 7 : Textes, Comms, Blocages"):
            in_high = st.text_area("Slide 1 - Highlights", "- Lancement réussi Market zone Nord\n- Record de commandes ce weekend")
            in_comms = st.text_area("Slide 5 - Comms & Brand (Social, RP, Interne, M+2)", "- Social : +12% abonnés, Top Post KFC\n- RP : 3 articles\n- Interne : Newsletter Riders\n- M+2 : Campagne Plages")
            in_bloc = st.text_area("Slide 6 - Blocages & Actions", "- Blocage: Manque coursiers vendredi (Pluie)\n- Action: Surge pricing activé\n- Blocage: Bug tech temporaire")
            in_road = st.text_area("Slide 7 - Priorités S+1", "- Lancement zones périphériques\n- Grosse campagne acquisition ciblée\n- Amélioration Batching")

    with tab3:
        st.write("### 📤 Export de la présentation")
        bg_image_css = ""
        if template_file:
            b64_img = encode_image(template_file)
            bg_image_css = f"background-image: url('{b64_img}'); background-size: cover; background-position: center;"
            
        def li_fmt(txt): return "".join([f"<li>{l.replace('- ', '').strip()}</li>" for l in txt.split('\n') if l.strip()])

        # Build Forecast HTML Table
        forecast_html = ""
        for _, row in edited_forecast.iterrows():
            if str(row['Restaurant']).strip():
                forecast_html += f"<tr><td style='font-weight:600;'>{row['Restaurant']}</td><td style='color:#F59E0B; font-weight:bold;'>{row['Glovo Orders (Estim.)']}</td><td style='color:#10B981; font-weight:bold;'>{row['Yassir Forecast']}</td></tr>"
        if not forecast_html: forecast_html = "<tr><td colspan='3'>Aucune prévision saisie.</td></tr>"

        # Build Top 10 Global HTML Table
        t10_html = ""
        for i, p in enumerate(kpis['top_partners']):
            t10_html += f"<tr><td>#{i+1} <b>{p['restaurant name'][:20]}</b></td><td>{p['orders']:,}</td><td>{p['gmv']:,.0f}</td></tr>"

        # Req vs Delivered
        req_del_pct = (kpis['orders'] / in_req * 100) if in_req > 0 else 0
        vs_budget = ((kpis['gmv'] - in_budget_gmv) / in_budget_gmv * 100) if in_budget_gmv > 0 else 0

        html_template = f"""<!DOCTYPE html>
        <html lang="fr">
        <head>
            <meta charset="UTF-8">
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;700;800&display=swap" rel="stylesheet">
            <style>
                body {{ background: #1E293B; font-family: 'Inter', sans-serif; display:flex; flex-direction:column; align-items:center; gap:40px; padding:40px; margin:0;}}
                .slide {{ width: 1280px; height: 720px; background-color: #FFFFFF; {bg_image_css} position: relative; overflow: hidden; box-shadow: 0 20px 40px rgba(0,0,0,0.5); border-radius:12px; }}
                .slide-content {{ padding: 50px 60px; height: 100%; display: flex; flex-direction: column; background: rgba(255,255,255,0.92); }}
                .header {{ display:flex; justify-content:space-between; border-bottom: 2px solid #E2E8F0; padding-bottom: 20px; margin-bottom:25px; }}
                .logo {{ font-size: 24px; font-weight:800; color:#0F172A;}} .logo span {{ color: #E11D48; }}
                .slide-tag {{ font-size: 16px; font-weight: 700; color: #64748B; text-transform: uppercase; }}
                h2.title {{ font-size: 32px; font-weight: 800; margin-bottom: 25px; color: #0F172A; }}
                h3 {{ color: #E11D48; font-size: 18px; margin: 0 0 15px 0; }}
                
                .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 30px; flex-grow:1; }}
                .grid-3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; }}
                .grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom:25px; }}
                
                .card {{ background: #F8FAFC; border: 1px solid #E2E8F0; padding: 25px; border-radius: 12px; }}
                .kpi-box {{ background: white; border-left: 4px solid #E11D48; padding: 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }}
                .kpi-lbl {{ font-size: 13px; color: #64748B; font-weight: 700; text-transform: uppercase; margin-bottom: 8px;}}
                .kpi-val {{ font-size: 32px; font-weight: 800; color: #0F172A; line-height:1.2;}}
                .kpi-sub {{ font-size: 13px; margin-top:5px; display:flex; justify-content:space-between;}}
                
                ul {{ list-style: none; padding: 0; margin:0;}}
                li {{ position: relative; padding-left: 18px; margin-bottom: 12px; font-size:16px; color:#1E293B; line-height:1.5;}}
                li::before {{ content: '•'; position: absolute; left: 0; color: #E11D48; font-weight: bold; font-size:20px; top:-2px;}}
                
                table {{ width: 100%; border-collapse: collapse; background: white; border-radius:8px; overflow:hidden; border:1px solid #E2E8F0; font-size:14px;}}
                th, td {{ padding: 12px 15px; text-align: left; border-bottom: 1px solid #E2E8F0; }}
                th {{ background: #F1F5F9; color: #475569; font-size:12px; text-transform:uppercase;}}
                
                .funnel {{ background: #F1F5F9; border-radius:8px; padding:15px; text-align:center; }}
                .funnel-step {{ background:#1E293B; color:white; padding:10px; margin-bottom:5px; border-radius:6px; font-weight:bold; display:flex; justify-content:space-between;}}
            </style>
        </head>
        <body>
            <!-- SLIDE 1 : EXECUTIVE SUMMARY -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">01. Executive Summary</div></div>
                <h2 class="title">Chiffres Clés ({start_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')})</h2>
                
                <div class="grid-4">
                    <div class="kpi-box"><div class="kpi-lbl">Commandes Livrées</div><div class="kpi-val">{kpis['orders']:,}</div>
                        <div class="kpi-sub">{get_trend(kpis['orders'], kpis_prev['orders'])} <span style="color:#64748B;">Vs Obj: {((kpis['orders']-15000)/15000*100):.1f}%</span></div></div>
                    <div class="kpi-box"><div class="kpi-lbl">GMV (MAD)</div><div class="kpi-val">{kpis['gmv']:,.0f}</div>
                        <div class="kpi-sub">{get_trend(kpis['gmv'], kpis_prev['gmv'])} <span style="color:#64748B;">Vs Bdgt: {vs_budget:+.1f}%</span></div></div>
                    <div class="kpi-box"><div class="kpi-lbl">Req vs Delivered</div><div class="kpi-val">{req_del_pct:.1f}%</div>
                        <div class="kpi-sub"><span style="color:#64748B;">Sur {in_req:,} requêtes</span></div></div>
                    <div class="kpi-box"><div class="kpi-lbl">Panier Moyen (AOV)</div><div class="kpi-val">{kpis['aov']:.1f}</div>
                        <div class="kpi-sub">{get_trend(kpis['aov'], kpis_prev['aov'])}</div></div>
                </div>
                
                <div class="grid-2">
                    <div class="card"><h3>🎯 Highlights de la semaine</h3><ul>{li_fmt(in_high)}</ul></div>
                    <div class="card"><h3>📊 Tendance & Comportement</h3><ul>
                        <li><strong>Fréquence:</strong> {in_freq} cmds / utilisateur.</li>
                        <li><strong>Répartition Hebdo:</strong> {in_weekdays}% Weekdays / {100-in_weekdays}% Weekend.</li>
                        <li><strong>Taux d'annulation global:</strong> {kpis['cancel_rate']:.1f}% {get_trend(kpis['cancel_rate'], kpis_prev['cancel_rate'], True, True)}</li>
                    </ul></div>
                </div>
            </div></div>

            <!-- SLIDE 2 : OPERATIONS -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">02. Opérations & Logistique</div></div>
                <div class="grid-2">
                    <div style="display:flex; flex-direction:column; gap:20px;">
                        <div class="card"><h3>⏱️ Délais de Livraison</h3>
                            <div style="font-size:36px; font-weight:800; margin-bottom:10px;">{kpis['avg_time']:.0f} min <span style="font-size:16px; color:#64748B; font-weight:normal;">Moyenne click-to-door</span></div>
                            <div style="display:flex; gap:15px; font-size:14px; font-weight:600; margin-bottom:15px;">
                                <div style="background:white; padding:8px 12px; border-radius:6px; border:1px solid #E2E8F0;">Food: <span style="color:#E11D48;">{in_t_food}m</span></div>
                                <div style="background:white; padding:8px 12px; border-radius:6px; border:1px solid #E2E8F0;">Groceries: <span style="color:#E11D48;">{in_t_groc}m</span></div>
                            </div>
                            <div style="font-size:14px; color:#475569;">Acceptation: <b>{in_t_acc}m</b> | Pickup: <b>{in_t_pick}m</b> | Trajet: <b>{in_t_traj}m</b></div>
                        </div>
                        <div class="card"><h3>🛵 Flotte & Support</h3><ul>
                            <li><strong>Coursiers Actifs:</strong> {in_fleet} | <strong>Mix 3PL:</strong> {in_mix}%</li>
                            <li><strong>Batching:</strong> {in_batch} | <strong>Auto-dispatch:</strong> {in_auto}%</li>
                            <li><strong>CPO:</strong> {in_cpo} MAD</li>
                            <li><strong>Support:</strong> {in_calls} appels | Remb: ~1.2% GMV</li>
                        </ul></div>
                    </div>
                    <div class="card"><h3>❌ Taux d'annulation & Top Raisons</h3>
                        <div style="font-size:36px; font-weight:800; color:#EF4444; margin-bottom:20px;">{kpis['cancel_rate']:.1f}%</div>
                        <ul>
                            {''.join([f"<li>{r[0]} ({r[1]} commandes)</li>" for r in kpis['cancel_reasons']])}
                        </ul>
                    </div>
                </div>
            </div></div>

            <!-- SLIDE 3 : SALES & AM -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">03. Sales & Account Management</div></div>
                <div class="grid-2" style="grid-template-columns: 1fr 1.3fr;">
                    <div style="display:flex; flex-direction:column; gap:20px;">
                        <div class="card"><h3>🤝 Pipeline & AM</h3><ul>
                            <li><strong>Nouveaux Signés:</strong> {in_signed}</li>
                            <li><strong>Onboardés:</strong> {in_onboarded}</li>
                            <li><strong>Highlight Deal:</strong> {in_topdeal}</li>
                            <li><strong>Taux de Churn:</strong> {in_churn}%</li>
                        </ul></div>
                        <div class="card" style="background:#1E293B; color:white; border:none;">
                            <h3 style="color:#F8FAFC;">⭐ Zoom Exclusivités (KOOL)</h3>
                            <div style="display:flex; justify-content:space-around; margin-top:20px;">
                                <div style="text-align:center;"><div style="font-size:28px; font-weight:800; color:#E11D48;">{in_kool}%</div><div style="font-size:12px; color:#94A3B8;">Poids GMV</div></div>
                                <div style="text-align:center;"><div style="font-size:28px; font-weight:800; color:#E11D48;">{in_kool_orders if 'in_kool_orders' in locals() else '38'}%</div><div style="font-size:12px; color:#94A3B8;">Poids Orders</div></div>
                            </div>
                        </div>
                    </div>
                    <div style="display:flex; flex-direction:column; gap:20px;">
                        <div class="card" style="padding:15px;"><h3>🥇 Top 10 Partenaires (GMV Global)</h3>
                            <table style="font-size:12px;"><tr><th>Restaurant</th><th>Orders</th><th>GMV</th></tr>{t10_html}</table>
                        </div>
                    </div>
                </div>
            </div></div>

            <!-- SLIDE 3B : FORECAST NOUVELLES SIGNATURES -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">03b. Orders Forecast (Nouvelles Signatures)</div></div>
                <div class="card" style="height:100%;">
                    <h3>📈 Top 15 Nouvelles Signatures (First Orders)</h3>
                    <table>
                        <thead><tr><th>Nouveau Partenaire</th><th>Orders Glovo (Estimés)</th><th>Forecast Yassir</th></tr></thead>
                        <tbody>{forecast_html}</tbody>
                    </table>
                </div>
            </div></div>

            <!-- SLIDE 4 : GROWTH & MARKETING -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">04. Growth & Marketing</div></div>
                <div class="grid-2">
                    <div class="card"><h3>📉 Funnel d'Acquisition App</h3>
                        <div class="funnel">
                            <div class="funnel-step" style="width:100%;"><span>Visites App</span><span>{in_visits:,}</span></div>
                            <div class="funnel-step" style="width:75%; margin:0 auto; background:#334155;"><span>Installs</span><span>{in_installs:,}</span></div>
                            <div class="funnel-step" style="width:50%; margin:0 auto; background:#E11D48;"><span>First Users</span><span>{in_ftu:,}</span></div>
                        </div>
                        <ul style="margin-top:20px;">
                            <li><strong>Conv. Visite -> Cmd:</strong> {(kpis['orders']/in_visits*100) if in_visits>0 else 0:.1f}%</li>
                            <li><strong>Coût Acquisition (CAC):</strong> {in_cac} MAD</li>
                        </ul>
                    </div>
                    <div class="card"><h3>🎟️ Performance Campagnes & Promos</h3>
                        <ul style="margin-bottom:20px;"><li><strong>Top Campagne:</strong> {in_campaign}</li></ul>
                        <div style="background:white; padding:20px; border-radius:8px; border:1px solid #E2E8F0;">
                            <div style="display:flex; justify-content:space-between; margin-bottom:15px; border-bottom:1px solid #E2E8F0; padding-bottom:10px;">
                                <strong>Budget Alloué (Expenses)</strong><span style="color:#EF4444; font-weight:800;">{in_budget_promo:,.0f} MAD</span>
                            </div>
                            <div style="display:flex; justify-content:space-between;">
                                <strong>GMV Promo Généré</strong><span style="color:#10B981; font-weight:800;">{in_gmv_promo:,.0f} MAD</span>
                            </div>
                        </div>
                        <div style="text-align:right; font-size:18px; font-weight:800; margin-top:15px; color:#1E293B;">
                            ROI Campagnes : <span style="color:#E11D48;">x{(in_gmv_promo/in_budget_promo) if in_budget_promo>0 else 0:.1f}</span>
                        </div>
                    </div>
                </div>
            </div></div>

            <!-- SLIDE 5 : COMMS & BRAND -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">05. Communication & Brand</div></div>
                <div class="card" style="height:100%;">
                    <h3>📢 Social, RP, Corporate & Calendrier M+2</h3>
                    <ul style="font-size:18px; line-height:1.8;">{li_fmt(in_comms)}</ul>
                </div>
            </div></div>

            <!-- SLIDE 6 & 7 : BLOCAGES & ROADMAP -->
            <div class="slide"><div class="slide-content">
                <div class="header"><div class="logo">LMD <span>Weekly</span></div><div class="slide-tag">06-07. Next Steps : Blocages & Roadmap S+1</div></div>
                <div class="grid-2">
                    <div class="card" style="background:#FEF2F2; border-color:#FCA5A5;">
                        <h3 style="color:#DC2626;">⚠️ Blocages, Frictions & Actions Correctives</h3>
                        <ul style="color:#7F1D1D;">{li_fmt(in_bloc)}</ul>
                    </div>
                    <div class="card" style="background:#F0FDF4; border-color:#86EFAC;">
                        <h3 style="color:#16A34A;">🚀 Top 3 Priorités Opérationnelles & Commerciales S+1</h3>
                        <ul style="color:#14532D;">{li_fmt(in_road)}</ul>
                    </div>
                </div>
            </div></div>

        </body>
        </html>
        """

        st.download_button("📥 Télécharger la présentation finale (HTML)", data=html_template, file_name=f"WBR_LMD_{start_date}.html", mime="text/html", use_container_width=True)
