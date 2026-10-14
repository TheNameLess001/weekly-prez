"""Consolidation conservative: provenance, granularité, complétude et N/D."""
from __future__ import annotations
import calendar
import hashlib
import json
import math
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
import pandas as pd

def norm(x):
    s = unicodedata.normalize('NFKD', str(x)).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+', '_', s).strip('_')

ALIASES = {
 'restaurant_id': ['restaurant id','rst id','id restaurant','store id','partner id','merchant id'],
 'restaurant_name': ['restaurant name','restaurant','nom restaurant','store name','partner name'],
 'city': ['city','ville'], 'brand': ['brand','enseigne','chain'],
 'order_id': ['order id','id commande','request id'], 'user_id': ['user id','customer id','id client','customer phone'],
 'date': ['date','created at','order date','date commande','day','jour','order day'],
 'start': ['period start','start date','date debut','week start'],
 'end': ['period end','end date','date fin','week end'],
 'status': ['status','statut','order status'],
 'requests': ['requests','received','commandes recues','recu'],
 'delivered': ['delivered orders','delivered','commandes livrees','livre'],
 'gmv': ['gmv','gmv delivered','gmv livre','item total'], 'ca': ['ca','revenue yassir','yassir revenue'],
 'commission': ['commission','commissions','restaurant commission'], 'commission_base': ['commission base','base commission'],
 'promo': ['promo cost','promotion cost','cout promotions','promo'],
 'benefit': ['benefit','contribution','final earnings'], 'users': ['unique users','users','utilisateurs uniques'],
 'delivered_users': ['delivered users','utilisateurs livres'],
 'new_users': ['new users','nouveaux utilisateurs'], 'is_new': ['is new','new customer'],
 'user_first_date': ['user first date','first order date','date premiere commande'],
 'repeat_users': ['repeat users','clients recurrents'],
 'cancelled': ['cancelled','canceled','annulations'],
 'cancel_restaurant': ['cancel restaurant','canc resto','annulations restaurant'],
 'cancel_driver': ['cancel driver','canc livreur','annulations livreur'],
 'cancel_client': ['cancel client','canc client','annulations client'],
 'cancel_other': ['cancel other','annulations administration'],
 'cancel_unknown': ['cancel unknown','annulations acteur inconnu'],
 'cancel_actor': ['cancel actor','cancellation actor','annulation acteur','cancelled by'],
 'cancel_reason': ['cancel reason','cancellation reason','motif annulation','cancellation reason'],
 'accept_min': ['accept min','acceptance minutes','temps acceptation'],
 'prep_min': ['prep min','preparation minutes','temps preparation'],
 'ready_min': ['ready min','ready minutes','preparation declaree'],
 'wait_min': ['wait min','driver wait minutes','attente livreur'],
 'travel_min': ['travel min','travel minutes','temps trajet'],
 'ctd_min': ['ctd min','click to door','delivery minutes','temps livraison','delivery time(M)'],
 'sla_met': ['sla met','on time','livraison a heure'], 'sla_n': ['sla n','sla orders'],
 'available_hours': ['available hours','heures disponibles'],
 'scheduled_hours': ['scheduled hours','heures prevues'],
 'active_days': ['active days','jours actifs'],
 'product_id': ['product id','item id','sku','id produit','id article'],
 'product_name': ['product name','item name','article','nom produit','nom article'],
 'category': ['category','categorie','product category'],
 'quantity': ['quantity','qty','quantite','units'],
 'food_item':['food item'], 'food_category':['food category'],
 'line_id': ['line id','order item id','id ligne'],
 'line_gmv': ['line gmv','item gmv','line total','montant ligne','gmv produit'],
}
TIMES = ['accept_min','prep_min','ready_min','wait_min','travel_min','ctd_min']
for t in TIMES:
    ALIASES[t+'_n'] = [t+'_n']
    ALIASES[t+'_median'] = [t+'_median']
STRINGS = {'restaurant_id','restaurant_name','city','brand','order_id','user_id','status','cancel_actor','cancel_reason','product_id','product_name','category','line_id','food_item','food_category'}
DATES = {'date','start','end','user_first_date'}
FIELDS = list(ALIASES)
ALIAS_LOOKUP = {norm(a): k for k,v in ALIASES.items() for a in [k]+v}

def number(x):
    if x is None or pd.isna(x) or str(x).strip() in ('','N/D','N/A','-'): return None
    if isinstance(x,(int,float)): return float(x) if math.isfinite(float(x)) else None
    s = str(x).strip().replace('\u202f','').replace('\xa0','').replace(' ','')
    s = re.sub(r'(?i)MAD|DH|EUR|%', '', s)
    if ',' in s and '.' in s:
        s = s.replace(',','') if s.rfind('.')>s.rfind(',') else s.replace('.','').replace(',','.')
    else: s = s.replace(',','.')
    try: return float(s)
    except ValueError: return None

def day(x):
    if x is None or pd.isna(x) or str(x).strip()=='': return None
    if isinstance(x,(int,float)):
        return (pd.Timestamp('1899-12-30')+pd.to_timedelta(float(x),unit='D')).date()
    try:
        # ISO and English month names do not pass through a day/month swap.
        s=str(x)
        return pd.to_datetime(x,dayfirst=not bool(re.match(r'^\d{4}-',s)),errors='raise').date()
    except (ValueError,TypeError,OverflowError): return None

def clean_id(x):
    if pd.isna(x) or str(x).strip()=='': return None
    return re.sub(r'^(\d+)\.0$',r'\1',str(x).strip())

def inspect_file(path):
    p=Path(path)
    if p.suffix.lower()=='.csv':
        try: frames={'CSV':pd.read_csv(p,sep=None,engine='python',dtype=str,encoding='utf-8-sig')}
        except UnicodeDecodeError: frames={'CSV':pd.read_csv(p,sep=None,engine='python',dtype=str,encoding='cp1252')}
    elif p.suffix.lower()=='.xlsx':
        frames=pd.read_excel(p,sheet_name=None,dtype=object,engine='openpyxl')
    else: raise ValueError('Format accepté : .xlsx ou .csv')
    out=[]
    for sheet,df in frames.items():
        cols=[str(c) for c in df.columns]
        mapping={}
        for c in cols:
            if norm(c) in ALIAS_LOOKUP and ALIAS_LOOKUP[norm(c)] not in mapping:
                mapping[ALIAS_LOOKUP[norm(c)]]=c
        from admin_earnings import is_admin
        admin=is_admin(cols,norm)
        if admin and 'gmv' in mapping:mapping['commission_base']=mapping['gmv']
        grain='orders' if 'order_id' in mapping else ('weekly' if re.search(r'(?:week|w)\s*\d{1,2}',p.stem,re.I) else 'daily')
        if 'date' not in mapping and 'start' not in mapping and 'requests' not in mapping and 'order_id' not in mapping: grain='registry'
        if ('product_id' in mapping or 'product_name' in mapping) and ('quantity' in mapping or 'line_gmv' in mapping): grain='items'
        demo=any(norm(c) in ('is_demo','demo_fictif') and any(str(v).strip().lower() in ('1','true','fictif','demo') for v in df[c].dropna()) for c in cols)
        out.append({'sheet':sheet,'columns':cols,'rows':len(df),'sample':df.head(4).astype(object).where(pd.notna(df.head(4)), '').astype(str).to_dict('records'),
                    'mapping':mapping,'profile':'admin_earnings' if admin else None,'grain':grain,'enabled':True,'complete':False,
                    'coverage_start':'','coverage_end':'','year':2026,'date_dayfirst':True,'demo':demo})
    return out

@dataclass
class Dataset:
    rows:list=field(default_factory=list)
    items:list=field(default_factory=list)
    registry:dict=field(default_factory=dict)
    quality:list=field(default_factory=list)
    coverage:list=field(default_factory=list)
    settings:dict=field(default_factory=dict)
    demo:bool=False

def ingest(sources,settings=None):
    ds=Dataset(settings=settings or {})
    seen={}; quarantined=set(); item_seen={}; item_conflicts=set()
    for src in sources:
        path=Path(src['path'])
        try:
            inspected=inspect_file(path)
            if path.suffix.lower()=='.xlsx': frames=pd.read_excel(path,sheet_name=None,dtype=object,engine='openpyxl')
            else:
                try: frames={'CSV':pd.read_csv(path,sep=None,engine='python',dtype=str,encoding='utf-8-sig')}
                except UnicodeDecodeError: frames={'CSV':pd.read_csv(path,sep=None,engine='python',dtype=str,encoding='cp1252')}
        except Exception as e:
            ds.quality.append({'type':'erreur fichier','source':path.name,'detail':str(e)}); continue
        ds.demo=ds.demo or bool(src.get('demo'))
        for auto in inspected:
            ds.demo=ds.demo or auto.get('demo',False)
            cfg=next((c for c in src.get('sheets',[]) if c['sheet']==auto['sheet']),auto)
            if not cfg.get('enabled',True): continue
            if auto.get('profile')=='admin_earnings' and not cfg.get('profile'):
                cfg['mapping']={**auto['mapping'],**cfg.get('mapping',{})}
                cfg['profile']='admin_earnings'
                if cfg.get('grain')=='registry':cfg['grain']='orders'
            sh=cfg['sheet']; grain=cfg.get('grain',auto['grain']); mp=cfg.get('mapping',auto['mapping'])
            admin=auto.get('profile')=='admin_earnings' and grain=='orders'
            if admin:
                from admin_earnings import DEFINITIONS
                for key,value in DEFINITIONS.items():
                    if not ds.settings.get(key):ds.settings[key]=value
            df=frames[sh]; label=f'{path.name} / {sh}'
            ds.quality.append({'type':'import','source':label,'detail':f'{len(df)} lignes, {grain}, {len(mp)} champs reconnus'})
            missing=[c for c in mp.values() if c not in df.columns]
            if missing: ds.quality.append({'type':'mapping invalide','source':label,'detail':str(missing)}); continue
            cs,ce=day(cfg.get('coverage_start')),day(cfg.get('coverage_end'))
            if cfg.get('complete') and cs and ce and cs<=ce:
                ds.coverage.append({'start':cs,'end':ce,'source':label,'grain':grain,
                                    'restaurant_ids':set(), 'all_partners':bool(cfg.get('all_partners',False))})
                cover=ds.coverage[-1]
            else: cover=None
            week_match=re.search(r'(?:week|w)\s*(\d{1,2})[_\s-]*(\d{4})?',path.stem,re.I)
            for idx,record in df.iterrows():
                r={'source':label,'line':int(idx)+2,'grain':grain,'complete':bool(cover)}
                for k,c in mp.items():
                    if k not in FIELDS: continue
                    v=record[c]
                    if k in STRINGS: r[k]=clean_id(v)
                    elif k in DATES:
                        if isinstance(v,str) and not re.match(r'^\d{4}-',v) and not cfg.get('date_dayfirst',True):
                            try: r[k]=pd.to_datetime(v,dayfirst=False,errors='raise').date()
                            except (ValueError,TypeError): r[k]=None
                        else: r[k]=day(v)
                    else:
                        r[k]=number(v)
                        if r[k] is None and not pd.isna(v) and str(v).strip():
                            ds.quality.append({'type':'valeur invalide','source':label,'detail':f'L{idx+2} {k}: {v}'})
                if admin:
                    from admin_earnings import derive
                    derive(record,r,norm,number)
                    r['profile']='admin_earnings'
                rid=r.get('restaurant_id')
                if not rid:
                    key=f"{norm(r.get('restaurant_name',''))}|{norm(r.get('city',''))}"
                    rid=ds.settings.get('verified_ids',{}).get(key)
                    if rid: r['restaurant_id']=str(rid)
                    else:
                        ds.quality.append({'type':'ID absent','source':label,'detail':f"L{idx+2}: {r.get('restaurant_name','?')}, clé à vérifier {key}"}); continue
                rid=str(rid); r['restaurant_id']=rid
                meta=ds.registry.setdefault(rid,{'id':rid,'name':r.get('restaurant_name') or rid,'city':r.get('city') or 'N/D','brand':r.get('brand') or ''})
                for k in ('restaurant_name','city','brand'):
                    target='name' if k=='restaurant_name' else k
                    if r.get(k) and meta[target] in ('','N/D',rid): meta[target]=r[k]
                if cover: cover['restaurant_ids'].add(rid)
                if grain=='registry': continue
                if grain=='items':
                    from merchandising import register_item
                    register_item(ds,r,item_seen,item_conflicts,cs,ce)
                    continue
                if grain=='orders':
                    r['start']=r['end']=r.get('date')
                    if not r.get('order_id'):
                        ds.quality.append({'type':'commande sans ID','source':label,'detail':f'L{idx+2} exclue'}); continue
                    key=('order',r['order_id'])
                else:
                    if grain=='daily': r['start']=r['end']=r.get('date') or r.get('start')
                    if grain=='weekly' and not r.get('start') and week_match:
                        try:
                            r['start']=date.fromisocalendar(int(week_match.group(2) or cfg.get('year',2026)),int(week_match.group(1)),1)
                            r['end']=r['start']+timedelta(days=6)
                        except ValueError: pass
                    key=('complement' if grain=='complement' else 'aggregate',rid,r.get('start'),r.get('end'))
                if not r.get('start') or not r.get('end') or r['end']<r['start']:
                    ds.quality.append({'type':'dates invalides','source':label,'detail':f'L{idx+2} exclue'}); continue
                if cover and not (cs<=r['start'] and r['end']<=ce):
                    ds.quality.append({'type':'hors couverture déclarée','source':label,'detail':f'L{idx+2} exclue'}); continue
                for t in TIMES:
                    if r.get(t) is not None and not 0<=r[t]<=1440:
                        r[t]=None
                        ds.quality.append({'type':'durée invalide','source':label,'detail':f'L{idx+2} {t}: durée exclue, commande conservée'})
                bad=[k for k,v in r.items() if k in FIELDS and k not in STRINGS|DATES|{'benefit','ca'} and v is not None and v<0]
                if bad or (r.get('requests') is not None and r.get('delivered') is not None and r['delivered']>r['requests']):
                    ds.quality.append({'type':'anomalie quantitative','source':label,'detail':f'L{idx+2} exclue : {bad or "Delivered > Requests"}'}); continue
                if key in quarantined: continue
                if key in seen:
                    prev=seen[key]
                    conflicts=[k for k in FIELDS if k not in ('restaurant_name','city','brand') and prev.get(k) is not None and r.get(k) is not None and prev[k]!=r[k]]
                    if conflicts:
                        ds.quality.append({'type':'conflit exclu','source':label,'detail':f'{key}: {conflicts}. Les deux versions sont exclues.'})
                        ds.rows.remove(prev); del seen[key]; quarantined.add(key)
                    else:
                        added=[k for k in FIELDS if prev.get(k) is None and r.get(k) is not None]
                        for k in added: prev[k]=r[k]
                        ds.quality.append({'type':'complément fusionné' if added else 'doublon exclu','source':label,'detail':f'{key}; champs: {added}'})
                        if label not in prev['source'].split(' + '): prev['source']+=' + '+label
                else: seen[key]=r; ds.rows.append(r)
    # Any conflicting row invalidates declared completeness of its source.
    damaged=set(q['source'] for q in ds.quality if q['type'] in ('ID absent','commande sans ID','dates invalides','anomalie quantitative','hors couverture déclarée','conflit exclu','article invalide','conflit article'))
    if quarantined:
        # A discarded previous version can belong to another source. Fail closed.
        ds.coverage=[]
    else: ds.coverage=[c for c in ds.coverage if c['source'] not in damaged]
    for rid in ds.registry:
        rows=[r for r in ds.rows if r['restaurant_id']==rid]
        orders=[r for r in rows if r['grain']=='orders']
        for a in list(rows):
            if a['grain'] in ('orders','complement'): continue
            overlaps=[r for r in orders if a['start']<=r['start']<=a['end']]
            if overlaps:
                ds.rows.remove(a)
                ds.coverage=[c for c in ds.coverage if c['source']!=a['source']]
                ds.quality.append({'type':'récapitulatif exclu','source':a['source'],'detail':f'{rid}, {a["start"]}–{a["end"]}: priorité aux commandes, sans cumul'})
        # Different aggregate intervals cannot safely overlap.
        aggs=[r for r in ds.rows if r['restaurant_id']==rid and r['grain']!='orders']
        blocked=set()
        for i,a in enumerate(aggs):
            for b in aggs[i+1:]:
                if (a['grain']=='complement') != (b['grain']=='complement'): continue
                if max(a['start'],b['start'])<=min(a['end'],b['end']): blocked.update((id(a),id(b)))
        for a in aggs:
            if id(a) in blocked:
                ds.rows.remove(a); ds.coverage=[c for c in ds.coverage if c['source']!=a['source']]
                ds.quality.append({'type':'chevauchement exclu','source':a['source'],'detail':f'{rid}: agrégats concurrents exclus, choisir une source'})
    from admin_earnings import extract_items
    explicit_item_orders={(item['restaurant_id'],item['order_id']) for item in ds.items}
    for r in ds.rows:
        if r.get('profile')=='admin_earnings' and (r['restaurant_id'],r['order_id']) not in explicit_item_orders:
            for item in extract_items(r,ds):
                from merchandising import register_item
                register_item(ds,item,item_seen,item_conflicts)
    # Recover ambiguous category only from an unambiguous source observation of the same product.
    categories={}
    for item in ds.items:
        if item.get('from_admin') and item['category']!='Catégorie non attribuable':
            categories.setdefault((item['restaurant_id'],item['product_id']),set()).add(item['category'])
    for item in ds.items:
        candidates=categories.get((item['restaurant_id'],item.get('product_id')),set())
        if item.get('from_admin') and item['category']=='Catégorie non attribuable' and len(candidates)==1:
            item['category']=next(iter(candidates));item['category_basis']='Catégorie unique observée pour ce produit dans une commande à catégorie unique'
    return ds

def days(a,b): return (b-a).days+1
def daterange(a,b): return {a+timedelta(days=i) for i in range(days(a,b))}
def periods(mode,anchor=None,current_start=None,current_end=None,previous_start=None,previous_end=None):
    anchor=day(anchor) or date.today()
    if mode in ('wow','wtd'):
        mon=anchor-timedelta(days=anchor.weekday())
        if mode=='wow': a,b=mon-timedelta(days=7),mon-timedelta(days=1)
        else: a,b=mon,anchor
        c,d=a-timedelta(days=7),b-timedelta(days=7)
    elif mode in ('mom','mtd'):
        first=anchor.replace(day=1); prev_end=first-timedelta(days=1); prev_start=prev_end.replace(day=1)
        if mode=='mom':
            a,b=prev_start,prev_end; d=a-timedelta(days=1); c=d.replace(day=1)
        else:
            a,b=first,anchor; c=prev_start; d=c+timedelta(days=min(anchor.day,prev_end.day)-1)
    else:
        a,b,c,d=map(day,[current_start,current_end,previous_start,previous_end])
    if not all([a,b,c,d]) or a>b or c>d: raise ValueError('Plages de dates invalides')
    return (a,b),(c,d)

LABELS={
 'requests':('Commandes reçues','count',1), 'delivered':('Commandes livrées','count',1),
 'sr':('Success Rate','rate',1),'gmv':('GMV livré','money',1),'aov':('Panier moyen','money',1),
 'ca':('CA Yassir','money',1),'commission':('Commissions Yassir','money',1),
 'commission_rate':('Commission effective','rate',1),'promo':('Coût des promotions','money',-1),
 'promo_per_order':('Promotion / livraison','money',-1),'benefit':('Contribution','money',1),
 'users':('Utilisateurs uniques reçus','count',1),'delivered_users':('Clients livrés uniques','count',1),
 'new_users':('Nouveaux clients livrés','count',1),
 'observed_new_users':('Premiers clients observés du restaurant','count',1),
 'repeat_users':('Clients avec ≥ 2 livraisons','count',1),'frequency':('Fréquence livrée','decimal',1),
 'repeat_share':('Clients récurrents','rate',1),'cancelled':('Annulations','count',-1),
 'cancel_rate':('Taux d’annulation','rate',-1),
 'cancel_other':('Annulations administration / autres','count',-1),
 'cancel_unknown':('Annulations acteur inconnu','count',-1),'sla_rate':('Respect SLA','rate',1),
 'availability':('Disponibilité','rate',1),'active_days':('Jours actifs','count',1),
 'requests_per_day':('Commandes reçues / jour','decimal',1),
 'delivered_per_day':('Livraisons / jour','decimal',1),'gmv_per_day':('GMV / jour','money',1),
}
TIME_LABELS={'accept_min':'Acceptation','prep_min':'Préparation estimée','ready_min':'Préparation déclarée','wait_min':'Attente après prêt','travel_min':'Trajet','ctd_min':'Click-to-door'}
for t,l in TIME_LABELS.items():
    feminine=t in ('accept_min','prep_min','ready_min','wait_min')
    LABELS[t]=(l+(' moyenne' if feminine else ' moyen'),'time',-1); LABELS[t+'_median']=(l+(' médiane' if feminine else ' médian'),'time',-1)
for actor,lab in [('restaurant','restaurant'),('driver','livreur'),('client','client'),('other','administration / autres'),('unknown','acteur inconnu')]:
    LABELS['cancel_'+actor]=('Annulations '+lab,'count',-1)
    LABELS['cancel_'+actor+'_rate']=('Taux annulations '+lab,'rate',-1)

def strict_sum(rows,key):
    vals=[r.get(key) for r in rows]
    return sum(vals) if vals and all(v is not None for v in vals) else None
def ratio(a,b): return a/b if a is not None and b is not None and b>0 else None
def known_status(r,settings):
    v=norm(r.get('status',''))
    custom=settings.get('statuses',{})
    if v in [norm(s) for s in custom.get('delivered',['delivered','livre','completed'])]: return 'delivered'
    if v in [norm(s) for s in custom.get('cancelled',['cancelled','canceled','annule','rejected','cancelled by user','cancelled by admin','restaurant rejected','cancelled by restaurant','cancelled by driver','cancelled by customer'])]: return 'cancelled'
    if v in [norm(s) for s in custom.get('other',['pending','accepted','preparing','ready','requested','en_cours'])]: return 'other'
    return None

def aggregate(ds,rid,interval):
    a,b=interval; requested=daterange(a,b)
    all_rows=[r for r in ds.rows if r['restaurant_id']==rid and r['start']<=b and r['end']>=a]
    complements=[r for r in all_rows if r['grain']=='complement']
    source_rows=[r for r in all_rows if r['grain']!='complement']
    crossing=[r for r in source_rows if r['start']<a or r['end']>b]
    rows=[r for r in source_rows if a<=r['start'] and r['end']<=b]
    covered=set()
    for c in ds.coverage:
        if c['grain'] in ('registry','complement','items'): continue
        if rid in c['restaurant_ids'] or c['all_partners']: covered|=daterange(c['start'],c['end']) & requested
    for r in rows:
        if r['grain']!='orders' and r['complete']: covered |= daterange(r['start'],r['end'])
    observed=set().union(*(daterange(r['start'],r['end']) for r in rows)) if rows else set()
    notes=[]; values={k:None for k in LABELS}; samples={}
    if crossing:
        notes.append('Agrégat à cheval sur les dates sélectionnées : aucun découpage ni prorata. Sélectionner les intervalles complets.')
        return {'values':values,'covered_days':len(covered),'observed_days':len(observed),'days':len(requested),'notes':notes,'samples':samples,'reasons':{},'trend':[]}
    exact=covered==requested
    if not exact: notes.append(f'Couverture non certifiée : {len(covered)}/{len(requested)} jours. Totaux observés seulement, absence de ligne ≠ zéro.')
    orders=[r for r in rows if r['grain']=='orders']; aggs=[r for r in rows if r['grain']!='orders']
    cancelled=[]; delivered=[]; reasons={}
    if orders:
        statuses=[known_status(r,ds.settings) for r in orders]
        if None in statuses: notes.append('Statuts inconnus ou absents : Delivered, SR et annulations N/D.')
        delivered=[r for r,s in zip(orders,statuses) if s=='delivered']
        cancelled=[r for r,s in zip(orders,statuses) if s=='cancelled']
        values['requests']=len(orders)
        if None not in statuses:
            values['delivered']=len(delivered); values['cancelled']=len(cancelled)
            for actor in ('restaurant','driver','client','other','unknown'):
                actor_map={'restaurant':['restaurant','resto','merchant'],'driver':['driver','livreur','courier'],'client':['client','customer','user'],'other':['other','autre']}
                values['cancel_'+actor]=sum((norm(r.get('cancel_actor','')) not in sum(actor_map.values(),[])) if actor=='unknown' else (norm(r.get('cancel_actor','')) in actor_map[actor]) for r in cancelled)
            if values['cancel_unknown']:
                notes.append(f"{values['cancel_unknown']} annulations sans acteur reconnu : ventilées en acteur inconnu, jamais imputées au livreur.")
            for r in cancelled:
                reason=r.get('cancel_reason') or 'Motif non renseigné'; reasons[reason]=reasons.get(reason,0)+1
        for key in ('gmv','ca','commission','commission_base','promo','benefit'):
            # By contract these imported amounts apply to delivered orders.
            if None not in statuses:
                values[key]=strict_sum(delivered,key) if delivered else (0. if any(key in r for r in orders) else None)
        if all(r.get('user_id') for r in orders): values['users']=len({r['user_id'] for r in orders})
        if None not in statuses and all(r.get('user_id') for r in delivered):
            users={r['user_id'] for r in delivered}; values['delivered_users']=len(users)
            counts=pd.Series([r['user_id'] for r in delivered]).value_counts()
            values['repeat_users']=int((counts>=2).sum())
            values['repeat_share']=ratio(values['repeat_users'],len(users))
            history=[r for r in ds.rows if r['restaurant_id']==rid and r['grain']=='orders' and r['start']<=b and known_status(r,ds.settings)=='delivered']
            if all(r.get('user_id') for r in history):
                first={}
                for r in history:first[r['user_id']]=min(first.get(r['user_id'],r['start']),r['start'])
                values['observed_new_users']=sum(a<=first[u]<=b for u in users)
                notes.append('Premiers clients observés : première livraison du restaurant dans les fichiers importés, historique antérieur non garanti. Nouveaux clients globaux distincts.')
            if delivered and all(r.get('user_first_date') for r in delivered):
                values['new_users']=len({r['user_id'] for r in delivered if a<=r['user_first_date']<=b})
            elif delivered and all(r.get('is_new') in (0,1) for r in delivered):
                values['new_users']=len({r['user_id'] for r in delivered if r['is_new']==1})
        for t in TIMES:
            vals=[r[t] for r in delivered if r.get(t) is not None]
            samples[t]=len(vals)
            if vals:
                values[t]=sum(vals)/len(vals); values[t+'_median']=float(pd.Series(vals).median())
        sla=[r['sla_met'] for r in delivered if r.get('sla_met') in (0,1)]
        threshold=number(ds.settings.get('sla_target_min'))
        if threshold is not None and threshold>0:
            sla=[r['sla_met'] if r.get('sla_met') in (0,1) else int(r['ctd_min']<=threshold) for r in delivered if r.get('sla_met') in (0,1) or r.get('ctd_min') is not None]
            notes.append(f'SLA calculé : click-to-door ≤ {threshold:g} min, seuil défini par l’utilisateur.')
        values['sla_rate']=sum(sla)/len(sla) if sla else None; samples['sla_rate']=len(sla)
        values['active_days']=len({r['start'] for r in delivered}) if None not in statuses else None
    if aggs:
        for key in ('requests','delivered','gmv','ca','commission','commission_base','promo','benefit','cancelled','cancel_restaurant','cancel_driver','cancel_client','cancel_other','cancel_unknown','active_days'):
            v=strict_sum(aggs,key)
            values[key]=(values[key]+v if values.get(key) is not None and v is not None else None) if orders else v
        # Unique users cannot be summed across intervals, even when disjoint.
        if not orders and len(aggs)==1:
            for key in ('users','delivered_users','new_users'): values[key]=aggs[0].get(key)
            values['repeat_users']=aggs[0].get('repeat_users')
            values['repeat_share']=ratio(aggs[0].get('repeat_users'),aggs[0].get('delivered_users'))
        else:
            for key in ('users','delivered_users','new_users','repeat_share'): values[key]=None
            notes.append('Identifiants clients absents sur les agrégats : uniques, nouveaux et récurrents multi-périodes N/D.')
        for t in TIMES:
            if orders or any(r.get(t) is None or r.get(t+'_n') is None for r in aggs): values[t]=None
            else:
                n=sum(r[t+'_n'] for r in aggs); values[t]=ratio(sum(r[t]*r[t+'_n'] for r in aggs),n); samples[t]=n
            values[t+'_median']=aggs[0].get(t+'_median') if not orders and len(aggs)==1 else None
        values['sla_rate']=ratio(strict_sum(aggs,'sla_met'),strict_sum(aggs,'sla_n')) if not orders else None
        samples['sla_rate']=strict_sum(aggs,'sla_n')
        values['availability']=ratio(strict_sum(aggs,'available_hours'),strict_sum(aggs,'scheduled_hours')) if not orders else None
    if not rows and exact:
        for key in ('requests','delivered','gmv','cancelled','cancel_restaurant','cancel_driver','cancel_client','users','delivered_users','new_users','active_days'): values[key]=0.
    if complements:
        cp=[r for r in complements if a<=r['start'] and r['end']<=b and r['complete']]
        cpdays=set().union(*(daterange(r['start'],r['end']) for r in cp)) if cp else set()
        if cpdays==requested:
            for k in ('gmv','ca','commission','commission_base','promo','benefit'):
                v=strict_sum(cp,k)
                if v is None: continue
                if values.get(k) is None: values[k]=v
                elif abs(values[k]-v)>max(.01,abs(v)*1e-8):
                    values[k]=None; notes.append(f'{k}: désaccord entre commandes et complément financier. KPI N/D.')
            notes.append('Complément financier sur intervalles complets : montants fusionnés, aucun cumul avec les commandes.')
        else: notes.append('Complément financier non exhaustif ou à cheval sur les dates : montants ignorés.')
    for k in ('ca','commission','promo','benefit'):
        if not ds.settings.get(k+'_definition'): values[k]=None
    values['sr']=ratio(values['delivered'],values['requests'])
    values['aov']=ratio(values['gmv'],values['delivered'])
    values['commission_rate']=ratio(values['commission'],values.get('commission_base')) if ds.settings.get('commission_base_definition') else None
    values['promo_per_order']=ratio(values['promo'],values['delivered'])
    values['frequency']=ratio(values['delivered'],values['delivered_users'])
    for actor in ('','restaurant','driver','client','other','unknown'):
        values['cancel_rate' if not actor else 'cancel_'+actor+'_rate']=ratio(values['cancelled' if not actor else 'cancel_'+actor],values['requests'])
    for k in ('requests','delivered','gmv'): values[k+'_per_day']=values[k]/len(requested) if values[k] is not None else None
    trend=[]
    for start in sorted({r['start'] for r in rows}):
        rr=[r for r in rows if r['start']==start]
        oo=[r for r in rr if r['grain']=='orders']
        dd=[r for r in oo if known_status(r,ds.settings)=='delivered']
        trend.append({'date':str(start),'delivered':len(dd) if oo and all(known_status(r,ds.settings) for r in oo) else strict_sum(rr,'delivered'),
                      'gmv':strict_sum(dd,'gmv') if dd else strict_sum(rr,'gmv'), 'grain':rr[0]['grain']})
    return {'values':values,'covered_days':len(covered),'observed_days':len(observed),'days':len(requested),'notes':notes,'samples':samples,'reasons':reasons,'trend':trend}

def fmt(value,kind='count',currency='MAD'):
    if value is None: return 'N/D'
    if kind=='rate': return f'{value*100:.1f} %'.replace('.',',')
    if kind=='time': return f'{(0.0 if abs(value)<.05 else value):.1f} min'.replace('.',',')
    if kind=='money': return f'{value:,.2f}'.replace(',',' ').replace('.',',')+' '+currency
    if kind=='decimal': return f'{value:.2f}'.replace('.',',')
    return f'{value:,.0f}'.replace(',',' ')

def compare(ds,rid,current,previous):
    cur,pre=aggregate(ds,rid,current),aggregate(ds,rid,previous)
    metrics=[]
    reliable=cur['covered_days']==cur['days'] and pre['covered_days']==pre['days']
    for k,(label,kind,direction) in LABELS.items():
        a,b=cur['values'].get(k),pre['values'].get(k)
        diff=a-b if a is not None and b is not None else None
        rel=diff/b if diff is not None and b>0 else None
        shown_diff=0. if diff is not None and kind=='rate' and abs(diff)<.0005 else diff
        delta=(f'{shown_diff*100:+.1f} pts'.replace('.',',') if kind=='rate' else (('+' if diff>0 else '')+fmt(diff,kind))) if diff is not None else 'N/A'
        growth=f'{(0 if abs(rel)<.0005 else rel)*100:+.1f} %'.replace('.',',') if rel is not None else ('Nouvelle activité' if b==0 and a is not None and a>0 and k in ('requests','delivered','gmv','ca','commission','users','delivered_users','new_users','observed_new_users','repeat_users','requests_per_day','delivered_per_day','gmv_per_day') else ('N/A (base négative)' if b is not None and b<0 else 'N/A'))
        neutral=diff is None or abs(diff)<({'rate':.0005,'time':.05,'money':.5,'decimal':.005}.get(kind,.5))
        reason=''
        if a is None or b is None:
            reason=('Référentiel de première commande globale absent' if k=='new_users' else
                    'Heures disponibles et heures prévues absentes' if k=='availability' else
                    'Statut SLA ou seuil SLA utilisateur absent' if k=='sla_rate' else
                    'Horodatages valides insuffisants sur au moins une période' if k.replace('_median','') in TIMES else
                    'Donnée, définition ou période absente dans la source')
        metrics.append({'key':k,'label':label,'kind':kind,'current':a,'previous':b,'delta':diff,'relative':rel,'delta_text':delta,'growth':growth,
                        'favorable':None if neutral else diff*direction>0,
                        'basis':'certifiée' if reliable else 'observée','unavailable_reason':reason,
                        'sample_current':cur['samples'].get(k.replace('_median','')),
                        'sample_previous':pre['samples'].get(k.replace('_median',''))})
    index={m['key']:m for m in metrics}
    facts=[]; actions=[]
    for k in ('delivered','gmv','sr'):
        m=index[k]; unit=m['delta_text'] if m['kind']=='rate' else m['growth']
        facts.append(f"{m['label']} : {fmt(m['current'],m['kind'])}, évolution {unit}.")
    candidates=[('sr','Conversion des commandes','Revoir les commandes non livrées et vérifier les principaux motifs avec le restaurant.','SR et taux d’annulation'),
                ('ctd_min','Délai de livraison','Examiner les commandes lentes et distinguer préparation, attente et trajet avant de décider.','Click-to-door et préparation'),
                ('gmv','Volume commercial','Comparer les jours et les créneaux les plus faibles, puis tester une mise en avant ciblée.','GMV et livraisons / jour'),
                ('promo_per_order','Coût promotionnel','Examiner le coût par livraison avant de reconduire la campagne.','Promotion / livraison'),
                ('repeat_share','Récurrence client','Tester une relance locale des clients déjà livrés et suivre leur retour.','Part des clients récurrents')]
    candidates.sort(key=lambda x: 0 if index[x[0]]['favorable'] is False else 1)
    for k,title,action,follow in candidates:
        m=index[k]
        if m['current'] is None or m['delta'] is None: continue
        actions.append({'title':title,'fact':f"{m['label']} : {fmt(m['current'],m['kind'])}, écart {m['delta_text']}.",'action':action,'follow':follow})
        if len(actions)==3: break
    if cur['values']['requests']==0 and cur['covered_days']==cur['days']:
        actions=[{'title':'Activation du partenaire','fact':f"0 commande reçue sur {cur['days']} jours certifiés.",'action':'Vérifier la publication du restaurant, les horaires et la possibilité de passer une commande.','follow':'Commandes reçues'},
                 {'title':'Disponibilité du catalogue','fact':f"0 livraison sur {cur['days']} jours certifiés.",'action':'Contrôler les produits commandables et les ruptures, puis tester une commande avec le restaurant.','follow':'Disponibilité et SR'},
                 {'title':'Première demande','fact':'Aucun utilisateur avec commande sur la période exhaustive.','action':'Si le restaurant est commandable, tester une mise en avant locale et mesurer les demandes générées.','follow':'Utilisateurs uniques et commandes reçues'}]
    missing=[m['label'] for m in metrics if m['current'] is None]
    fallbacks=[]
    if cur['covered_days']<cur['days']:
        fallbacks.append({'title':'Dates et exhaustivité','fact':f"Couverture certifiée {cur['covered_days']}/{cur['days']} jours.",'action':'Vérifier le périmètre et les jours absents. Les variations affichées portent sur les données importées.','follow':'Jours certifiés et écarts calculables'})
    if cur['values']['users'] is None or cur['values']['new_users'] is None:
        fallbacks.append({'title':'Mesure des clients','fact':'Uniques ou nouveaux utilisateurs N/D sur la période.','action':'Récupérer les User ID et la date de première commande issue du référentiel clients.','follow':'Uniques, nouveaux et fréquence'})
    financial=[k for k in ('commission','promo','benefit') if cur['values'][k] is None]
    if financial:
        fallbacks.append({'title':'Mesure économique','fact':f"{len(financial)} mesures absentes : "+', '.join(financial)+'.','action':'Valider les définitions et importer les montants du complément financier sur les mêmes dates.','follow':'Commission, promotion et contribution'})
    missing_times=[k for k in TIMES if cur['values'][k] is None]
    if missing_times:
        fallbacks.append({'title':'Mesure des délais','fact':f"{len(missing_times)}/{len(TIMES)} durées opérationnelles N/D.",'action':'Compléter les observations de durée ou leurs effectifs avant de fixer une priorité opérationnelle.','follow':'Effectifs des durées et click-to-door'})
    for action in fallbacks:
        if len(actions)>=3: break
        actions.append(action)
    if cur['days']!=pre['days']:
        facts.append('Durées différentes : comparer aussi les volumes par jour calendaire sélectionné.')
    from merchandising import compare_products
    merch=compare_products(ds,rid,current,previous)
    if merch['drops']:
        m=merch['drops'][0]
        actions=[{'title':'Produit en baisse','fact':f"{m['name'][:40]} : {m['growth']} en unités/jour.",
                  'action':'Vérifier disponibilité, prix et visibilité de ce produit avant de tester une relance ciblée.',
                  'follow':'Unités/jour de '+m['name'][:40]},*actions[:2]]
    return {'restaurant':ds.registry[rid],'current_dates':[str(x) for x in current],'previous_dates':[str(x) for x in previous],
            'current':cur,'previous':pre,'metrics':metrics,'facts':facts,'actions':actions[:3],
            'demo':ds.demo,'definitions':ds.settings,'sources':sorted({r['source'] for r in ds.rows if r['restaurant_id']==rid}),
            'comparison_basis':'certifiée' if reliable else 'observée',
            'time_definitions':__import__('admin_earnings').FORMULAS,
            'quality_count':len([q for q in ds.quality if q['type']!='import']), 'merchandising':merch}

def serializable(x):
    if isinstance(x,(date,Path)): return str(x)
    if isinstance(x,set): return sorted(x)
    raise TypeError(type(x).__name__)

def dumps(x): return json.dumps(x,ensure_ascii=False,default=serializable,allow_nan=False)
