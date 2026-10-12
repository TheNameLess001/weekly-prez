"""Online Streamlit UI for the validated Yassir reporting engine."""
from __future__ import annotations
import hashlib
import hmac
import json
import os
import sys
from datetime import date
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
import pandas as pd
import streamlit as st
from engine import FIELDS,day,periods,compare,fmt,dumps
from deck import pdf_converter
from online_workspace import Workspace,DEFINITION_KEYS,THEME_COLORS

st.set_page_config(page_title='Yassir Performance Studio',page_icon='📊',layout='wide')
st.markdown('''<style>
.block-container{max-width:1450px;padding-top:2rem}
[data-testid="stMetric"]{background:white;border:1px solid #e6deef;border-radius:14px;padding:18px}
h1,h2,h3{color:#24103F!important}.hero{background:linear-gradient(110deg,#24103F,#471b79);padding:25px 30px;border-radius:18px;color:white;margin-bottom:20px}.hero h1{color:white!important;font-size:31px;margin:0}.hero p{margin:8px 0 0;color:#e0d0f3}
</style>''',unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>Yassir Performance Studio</h1><p>Comparer le portefeuille. Comprendre les ventes. Générer les présentations.</p></div>',unsafe_allow_html=True)

def app_password():
    value=os.environ.get('APP_PASSWORD','')
    if not value:
        try:value=str(st.secrets.get('APP_PASSWORD',''))
        except FileNotFoundError:value=''
    if value=='REMPLACER_PAR_UN_SECRET_LONG_UNIQUE':return ''
    return value

password=app_password();protected=bool(password)
if protected and not st.session_state.get('authenticated',False):
    st.subheader('Connexion à l’espace équipe')
    with st.form('login'):
        candidate=st.text_input('Mot de passe',type='password',key='login_password')
        submitted=st.form_submit_button('Se connecter',type='primary')
    if submitted:
        if hmac.compare_digest(candidate.encode('utf-8'),password.encode('utf-8')):
            st.session_state['authenticated']=True;st.session_state.pop('login_password',None);st.rerun()
        else:st.error('Mot de passe incorrect.')
    st.stop()

if 'workspace' not in st.session_state:st.session_state.workspace=Workspace()
ws=st.session_state.workspace
if 'restore_selection' in st.session_state and ws.dataset.registry:
    st.session_state['selected_ids']=[rid for rid in st.session_state.pop('restore_selection') if rid in ws.dataset.registry]
if 'restore_period' in st.session_state:
    saved=st.session_state.pop('restore_period')
    if saved.get('mode') in ('WoW · dernière semaine complète','MoM · dernier mois complet','Semaine courante · même durée','Mois courant · même durée','Deux périodes libres'):
        st.session_state['period_mode']=saved['mode']
    for key in ('anchor','current_start','current_end','previous_start','previous_end'):
        if day(saved.get(key)):st.session_state[key]=day(saved[key])
with st.sidebar:
    st.subheader('Votre espace')
    if not protected:st.info('Mode démonstration. Configurez APP_PASSWORD dans les Secrets pour activer les imports réels.')
    else:st.caption('Espace équipe connecté')
    if st.button('Charger la démo Belhaj',key='load_demo',type='primary'):
        ws.load_demo();st.session_state['selected_ids']=['DEMO_BELHAJ'];st.session_state['period_mode']='Deux périodes libres';st.session_state['last_preview']=None;st.rerun()
    if st.button('Vider ma session',key='clear_session'):
        ws.reset();st.session_state['selected_ids']=[];st.session_state['last_preview']=None;st.rerun()
    if protected and st.button('Se déconnecter'):
        ws.close();st.session_state.clear();st.rerun()
    st.caption('Les fichiers sont traités sur le serveur et isolés par session. Une fermeture ou un redémarrage peut effacer les données temporaires. Téléchargez vos livrables et préférences.')

if ws.dataset.demo:st.warning('DÉMONSTRATION FICTIVE — aucun chiffre ni produit ne représente une performance réelle de Belhaj.')
import_tab,quality_tab,selection_tab,preview_tab,export_tab,preferences_tab=st.tabs(['1 · Importer','2 · Vérifier','3 · Sélectionner','4 · Analyser','5 · Générer','Préférences'])

with import_tab:
    st.subheader('Importer vos fichiers')
    st.caption('Excel .xlsx et CSV, plusieurs fichiers et toutes les feuilles. Commandes, portefeuille, finances et lignes articles sont consolidés séparément.')
    uploaded=st.file_uploader('Fichiers de performance',type=['csv','xlsx'],accept_multiple_files=True,disabled=not protected,key='data_upload')
    if st.button('Importer les fichiers',disabled=not protected or not uploaded,key='import_files',type='primary'):
        with st.spinner('Inspection et consolidation…'):errors,n=ws.add_uploads([(f.name,f.getvalue()) for f in uploaded])
        for error in errors:st.error(error)
        if n:st.success(f'{n} fichier(s) ajouté(s). Vérifiez le mapping et les dates.')
    if ws.sources:
        src_id=st.selectbox('Source à configurer',[s['id'] for s in ws.sources],format_func=lambda v:next(s['name'] for s in ws.sources if s['id']==v),key='source_selector')
        src=next(s for s in ws.sources if s['id']==src_id)
        sheet_name=st.selectbox('Feuille',[x['sheet'] for x in src['sheets']],key=f'sheet_{src_id}')
        sheet=next(x for x in src['sheets'] if x['sheet']==sheet_name)
        st.caption(f"{src['name']} / {sheet['sheet']} — {sheet['rows']} lignes")
        with st.expander('Échantillon original',expanded=False):st.dataframe(pd.DataFrame(sheet['sample']),hide_index=True)
        key=src_id+'_'+hashlib.sha256(sheet_name.encode()).hexdigest()[:8]+f'_{ws.revision}'
        grains={'orders':'Commande','items':'Lignes articles','daily':'Jour','weekly':'Semaine','period':'Intervalle','registry':'Référentiel','complement':'Complément financier'}
        with st.form('mapping_'+key):
            c1,c2,c3=st.columns(3)
            enabled=c1.checkbox('Importer cette feuille',value=sheet.get('enabled',True))
            grain=c2.selectbox('Grain',list(grains),index=list(grains).index(sheet['grain']),format_func=grains.get)
            dayfirst=c3.checkbox('Dates jour/mois',value=sheet.get('date_dayfirst',True))
            complete=st.checkbox('Source exhaustive sur ces dates',value=sheet.get('complete',False),help='Certifier uniquement après vérification. Les articles exigent leur propre certification.')
            c1,c2,c3=st.columns(3)
            start=c1.date_input('Début couvert',value=day(sheet.get('coverage_start')) or None)
            end=c2.date_input('Fin couverte',value=day(sheet.get('coverage_end')) or None)
            year=c3.number_input('Année ISO si absente du nom',min_value=2000,max_value=2100,value=int(sheet.get('year',2026)))
            all_partners=st.checkbox('La source couvre tous les partenaires importés',value=sheet.get('all_partners',False),help='Permet de distinguer zéro activité et absence de données du portefeuille.')
            with st.expander('Mapping des colonnes',expanded=False):
                mapping={};cols=st.columns(3)
                for i,field in enumerate(FIELDS):
                    options=['Non disponible',*sheet['columns']];existing=sheet.get('mapping',{}).get(field)
                    value=cols[i%3].selectbox(field,options,index=options.index(existing) if existing in options else 0,key='map_'+key+field)
                    if value!='Non disponible':mapping[field]=value
            applied=st.form_submit_button('Appliquer le mapping et les dates',type='primary')
        if applied:
            if complete and (not start or not end or start>end):st.error('Certification : saisir un intervalle valide.')
            else:
                sheet.update(enabled=enabled,grain=grain,date_dayfirst=dayfirst,complete=complete,coverage_start=str(start) if start else '',coverage_end=str(end) if end else '',year=int(year),all_partners=all_partners,mapping=mapping)
                ws.rebuild();st.session_state['last_preview']=None;st.rerun()
        if st.button('Retirer cette source',key='remove_'+src_id):ws.remove_source(src_id);st.session_state['last_preview']=None;st.rerun()
    else:st.info('Importez vos fichiers ou chargez la démonstration depuis le menu de gauche.')

with quality_tab:
    st.subheader('Contrôle de qualité')
    cols=st.columns(4)
    for col,label,value in zip(cols,['Fichiers','Commandes / agrégats','Lignes articles','Restaurants'],[len(ws.sources),len(ws.dataset.rows),len(ws.dataset.items),len(ws.dataset.registry)]):col.metric(label,value)
    q=ws.dataset.quality
    if q:st.dataframe(pd.DataFrame(q),hide_index=True,width='stretch')
    st.download_button('Télécharger le contrôle JSON',dumps({'quality':q,'coverage':ws.dataset.coverage}).encode(),file_name='controle_qualite.json',mime='application/json',key='quality_download')
    with st.expander('Correspondances Restaurant ID vérifiées'):
        verified=st.text_area('JSON nom_normalisé|ville_normalisée → ID vérifié',value=json.dumps(ws.settings.get('verified_ids',{}),ensure_ascii=False,indent=2),key=f'verified_{ws.revision}')
        if st.button('Enregistrer les correspondances'):
            try:
                values=json.loads(verified)
                if not isinstance(values,dict):raise ValueError('Un objet JSON est nécessaire')
                ws.settings['verified_ids']={str(k):str(v) for k,v in values.items()};ws.rebuild();st.rerun()
            except ValueError as e:st.error(str(e))

with selection_tab:
    st.subheader('Votre portefeuille')
    restaurants=list(ws.dataset.registry.values());by_id={r['id']:r for r in restaurants}
    c1,c2,c3=st.columns(3)
    search=c1.text_input('Rechercher un nom ou ID',key='search')
    city=c2.selectbox('Ville',['Toutes',*sorted({r['city'] for r in restaurants})],key=f'city_{ws.revision}')
    brand=c3.selectbox('Enseigne',['Toutes',*sorted({r['brand'] for r in restaurants if r['brand']})],key=f'brand_{ws.revision}')
    filtered=[r for r in restaurants if search.lower() in (r['name']+' '+r['id']).lower() and (city=='Toutes' or r['city']==city) and (brand=='Toutes' or r['brand']==brand)]
    st.session_state['selected_ids']=[rid for rid in st.session_state.get('selected_ids',[]) if rid in by_id]
    c1,c2=st.columns(2)
    if c1.button('Sélectionner tous les résultats'):st.session_state['selected_ids']=list(dict.fromkeys([*st.session_state['selected_ids'],*[r['id'] for r in filtered]]))
    if c2.button('Tout désélectionner'):st.session_state['selected_ids']=[]
    selected=st.multiselect('Restaurants à exporter',list(by_id),format_func=lambda rid:by_id[rid]['name']+' · '+rid,key='selected_ids')
    st.caption(f'{len(selected)} restaurant(s) sélectionné(s). La recherche et les filtres contrôlent le tableau et le bouton « Sélectionner tous les résultats ».')
    if filtered:st.dataframe(pd.DataFrame(filtered).rename(columns={'id':'Restaurant ID','name':'Restaurant','city':'Ville','brand':'Enseigne'}),hide_index=True,width='stretch')

with preview_tab:
    st.subheader('Comparaison des périodes')
    modes={'WoW · dernière semaine complète':'wow','MoM · dernier mois complet':'mom','Semaine courante · même durée':'wtd','Mois courant · même durée':'mtd','Deux périodes libres':'custom'}
    c1,c2=st.columns(2)
    mode=c1.selectbox('Mode de comparaison',list(modes),key='period_mode')
    anchor=c2.date_input('Date de référence',value=date.today(),key='anchor')
    if modes[mode]=='custom':
        c1,c2,c3,c4=st.columns(4)
        a=c1.date_input('Actuel début',value=date(2026,9,28),key='current_start');b=c2.date_input('Actuel fin',value=date(2026,10,4),key='current_end')
        c=c3.date_input('Référence début',value=date(2026,9,21),key='previous_start');d=c4.date_input('Référence fin',value=date(2026,9,27),key='previous_end')
    else:a=b=c=d=None
    try:
        current,previous=periods(modes[mode],anchor,a,b,c,d);valid_period=True
        period_config={'mode':mode,'anchor':str(anchor),'current_start':str(current[0]),'current_end':str(current[1]),'previous_start':str(previous[0]),'previous_end':str(previous[1])}
        st.caption(f'Actuel : {current[0]} au {current[1]} ({(current[1]-current[0]).days+1} jours). Référence : {previous[0]} au {previous[1]} ({(previous[1]-previous[0]).days+1} jours).')
    except ValueError as e:st.error(str(e));valid_period=False;period_config={}
    if selected and valid_period:
        rid=st.selectbox('Restaurant à analyser',selected,format_func=lambda r:by_id[r]['name'],key='preview_restaurant')
        report=compare(ws.dataset,rid,current,previous)
        notes=report['current']['notes']+report['previous']['notes']
        for note in dict.fromkeys(notes):st.warning(note)
        metrics={m['key']:m for m in report['metrics']};cols=st.columns(6)
        for col,key in zip(cols,['requests','delivered','sr','gmv','aov','users']):
            m=metrics[key];col.metric(m['label'],fmt(m['current'],m['kind'],ws.theme['currency']),m['delta_text'] if m['kind']=='rate' else m['growth'],delta_color='normal' if m['favorable'] is not None else 'off')
        st.markdown('**Constats observés**')
        for fact in report['facts']:st.write(fact)
        c1,c2=st.columns(2)
        if report['current']['trend']:
            chart=pd.DataFrame(report['current']['trend']).set_index('date')
            c1.subheader('Livraisons');c1.line_chart(chart[['delivered']],color='#6C2BD9')
            c2.subheader('GMV livré');c2.line_chart(chart[['gmv']],color='#6C2BD9')
        merch=report['merchandising'];st.subheader('Ventes produits, catégories et Top 3')
        for note in dict.fromkeys(merch['current']['notes']+merch['previous']['notes']):st.caption(note)
        cards=st.columns(3)
        for i,r in enumerate(merch['top3']):cards[i].metric(f"#{i+1} · {r['name']}",fmt(r['current_qty'])+' unités',r['growth'])
        if merch['products']:
            rows=[{'Produit':r['name'],'Catégorie':r['category'],'Unités actuelles':r['current_qty'],'Unités réf.':r['previous_qty'],'GMV produit':r['current_gmv'],'Part unités':r['share'],'Évolution /jour':r['growth']} for r in merch['products']]
            st.dataframe(pd.DataFrame(rows),hide_index=True,width='stretch')
        if merch['categories']:
            st.bar_chart(pd.DataFrame([{'Catégorie':r['name'],'Unités actuelles':r['current_qty'],'Unités référence':r['previous_qty']} for r in merch['categories']]).set_index('Catégorie'),color=['#6C2BD9','#B8A4D3'])
        c1,c2=st.columns(2)
        for col,title,rows in [(c1,'Top hausses',merch['risers']),(c2,'Drop produits',merch['drops'])]:
            col.subheader(title)
            if rows:col.dataframe(pd.DataFrame([{'Produit':r['name'],'Écart unités/jour':r['delta_per_day'],'Évolution':r['growth']} for r in rows]),hide_index=True)
            else:col.caption('Aucune variation mesurable.' if merch['reliable'] else 'N/A — périodes articles non certifiées.')
        with st.expander('Tous les KPI'):
            st.dataframe(pd.DataFrame([{'KPI':m['label'],'Actuel':fmt(m['current'],m['kind'],ws.theme['currency']),'Référence':fmt(m['previous'],m['kind'],ws.theme['currency']),'Écart':m['delta_text'],'Variation':m['growth']} for m in report['metrics']]),hide_index=True,width='stretch')
        st.download_button('Télécharger le diagnostic JSON',dumps(report).encode(),file_name='diagnostic_'+rid+'.json',mime='application/json')
    else:st.info('Sélectionnez un restaurant pour voir les résultats.')

def run_export(ids):
    progress=st.progress(0,text='Préparation du lot…');live=st.empty()
    def tick(n,total,item):
        progress.progress(n/total,text=f'{n}/{total} restaurants exportés');live.caption(('Terminé : ' if item['ok'] else 'Échec : ')+item['id'])
    with st.spinner('Génération des présentations…'):result=ws.generate(ids,current,previous,pdf,tick)
    progress.progress(1.0,text='Lot terminé');return result

with export_tab:
    st.subheader('Générer les présentations')
    st.write('14 slides par restaurant : performance, ventes produits, catégories, Top 3, hausses/Drop, plan d’action et annexes.')
    pdf=st.checkbox('Ajouter les PDF',value=bool(pdf_converter()),disabled=not bool(pdf_converter()),key='pdf_export')
    if not pdf_converter():st.info('LibreOffice indisponible sur ce serveur. PPTX éditables et ZIP restent opérationnels.')
    st.caption(f'{len(selected)} restaurant(s) sélectionné(s). Les livrables sont temporaires : téléchargez le lot après génération.')
    if st.button('Générer la sélection',key='generate',type='primary',disabled=not selected or not valid_period):run_export(selected)
    for i,result in enumerate(reversed(ws.results)):
        st.divider();st.markdown('**Lot disponible**')
        archive=Path(result['zip']);st.download_button('Télécharger le ZIP du lot',archive.read_bytes(),file_name=archive.name,mime='application/zip',key=f'zip_{i}',type='primary')
        failed=[item['id'] for item in result['results'] if not item['ok']]
        if failed and st.button(f'Relancer {len(failed)} export(s) échoué(s)',key=f'retry_{i}'):
            last=result['request'];retry=ws.generate(failed,last['current'],last['previous'],last['pdf']);st.rerun()
        for j,item in enumerate(result['results']):
            if not item['ok']:st.error(item['id']+' : '+item['error']);continue
            with st.expander(by_id.get(item['id'],{}).get('name',item['id'])):
                c1,c2=st.columns(2)
                for col,kind in ((c1,'pptx'),(c2,'pdf')):
                    if item.get(kind):
                        path=Path(item[kind]);col.download_button('Télécharger '+kind.upper(),path.read_bytes(),file_name=path.name,key=f'{kind}_{i}_{j}')
                if item.get('warning'):st.warning(item['warning'])
        with st.expander('Journal du lot'):st.dataframe(pd.DataFrame(result['results']).drop(columns=['pptx','pdf'],errors='ignore'),hide_index=True)

with preferences_tab:
    st.subheader('Définitions financières et préférences')
    with st.form('definitions_'+str(ws.revision)):
        values={};cols=st.columns(2)
        labels={'ca_definition':'CA Yassir','commission_definition':'Commissions','commission_base_definition':'Base de commission','promo_definition':'Coût promotionnel','benefit_definition':'Contribution / Benefit','line_gmv_definition':'Montant ligne produit','tax_basis':'Convention HT/TTC et devise'}
        for i,k in enumerate(DEFINITION_KEYS):values[k]=cols[i%2].text_area(labels[k],value=ws.settings.get(k,''),height=90)
        statuses=st.text_area('Statuts personnalisés (JSON facultatif)',value=json.dumps(ws.settings.get('statuses',{}),ensure_ascii=False),height=90)
        submitted=st.form_submit_button('Enregistrer les définitions',type='primary')
    if submitted:
        try:
            mapping=json.loads(statuses)
            if not isinstance(mapping,dict) or any(not isinstance(v,list) for v in mapping.values()):raise ValueError('Statuts : objet JSON avec des listes')
            ws.settings.update(values)
            if mapping:ws.settings['statuses']=mapping
            else:ws.settings.pop('statuses',None)
            ws.rebuild();st.rerun()
        except ValueError as e:st.error(str(e))
    with st.expander('Thème et logos'):
        with st.form('theme_form'):
            c1,c2,c3=st.columns(3);palette={}
            for i,k in enumerate(THEME_COLORS):palette[k]=[c1,c2,c3][i%3].color_picker(k,value=ws.theme[k],key=f'color_{k}_{ws.revision}')
            font=c1.text_input('Police PowerPoint',value=ws.theme['font']);currency=c2.text_input('Devise',value=ws.theme['currency'])
            if st.form_submit_button('Appliquer le thème'):
                try:ws.theme.update(Workspace.clean_theme({**palette,'font':font,'currency':currency}));ws.results=[];st.rerun()
                except ValueError as e:st.error(str(e))
        targets=['Yassir',*by_id];target=st.selectbox('Logo à remplacer',targets,format_func=lambda rid:'Yassir' if rid=='Yassir' else by_id[rid]['name'])
        logo=st.file_uploader('Logo PNG, JPG ou SVG',type=['png','jpg','jpeg','svg'],key='logo_upload',disabled=not protected)
        if st.button('Importer le logo',disabled=not protected or not logo):
            try:path=ws.save_logo(logo.name,logo.getvalue(),None if target=='Yassir' else target);st.image(str(path),width=200);st.success('Logo enregistré pour cette session.')
            except Exception as e:st.error(str(e))
    st.subheader('Sauvegarder et restaurer vos réglages')
    st.caption('Le JSON contient les mappings, définitions, thème, sélection et périodes. Il ne contient ni commandes, ni logos, ni mot de passe. Sur un fichier modifié, les dates doivent être recertifiées.')
    st.download_button('Télécharger mes préférences',ws.config_bytes(selected,period_config),file_name='yassir_preferences.json',mime='application/json',key='config_download')
    config=st.file_uploader('Préférences JSON',type=['json'],key='config_upload',disabled=not protected)
    if st.button('Restaurer les préférences',disabled=not protected or not config):
        try:
            obj=ws.apply_config(config.getvalue())
            st.session_state['restore_selection']=obj.get('selection',[]);st.session_state['restore_period']=obj.get('period',{});st.success('Préférences restaurées. Les mappings seront réutilisés lors de l’import des mêmes fichiers.')
            st.rerun()
        except (ValueError,TypeError) as e:st.error(str(e))
