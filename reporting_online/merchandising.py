"""Delivered item lines, independent of order-level GMV and order counts."""
from collections import defaultdict

def register_item(ds,r,seen,conflicts,cs=None,ce=None):
    from engine import norm
    reason=None
    if not r.get('date'): reason='Date article absente ou invalide'
    elif not r.get('order_id') or not r.get('line_id'): reason='Order ID et Line ID nécessaires pour dédupliquer les lignes articles'
    elif not r.get('product_id') and not r.get('product_name'): reason='Produit non identifié'
    elif any(r.get(k) is not None and r[k]<0 for k in ('quantity','line_gmv')): reason='Quantité ou montant ligne négatif'
    elif r.get('quantity') is None and r.get('line_gmv') is None: reason='Quantité et montant ligne absents'
    elif r['complete'] and not cs<=r['date']<=ce: reason='Ligne hors couverture déclarée'
    if reason:
        ds.quality.append({'type':'article invalide','source':r['source'],'detail':f"L{r['line']}: {reason}"}); return
    # A name-only match is explicitly exact and scoped to its category and restaurant.
    r['product_key']=r.get('product_id') or 'NAME:'+norm(r['product_name'])+'|'+norm(r.get('category') or 'Non catégorisé')
    r['start']=r['end']=r['date']
    key=(r['restaurant_id'],r['order_id'],r['line_id'])
    if key in conflicts: return
    if key in seen:
        old=seen[key]
        clash=[k for k in ('product_key','category','quantity','line_gmv','date','status') if old.get(k) is not None and r.get(k) is not None and old[k]!=r[k]]
        if clash:
            ds.items.remove(old); del seen[key]; conflicts.add(key)
            for source in {old['source'],r['source']}:
                ds.quality.append({'type':'conflit article','source':source,'detail':f'{key}: {clash}. Versions exclues.'})
        else:
            added=[k for k,v in r.items() if old.get(k) is None and v is not None]
            for k in added: old[k]=r[k]
            ds.quality.append({'type':'complément article' if added else 'doublon article exclu','source':r['source'],'detail':str(key)})
    else: seen[key]=r;ds.items.append(r)

def item_period(ds,rid,interval):
    from engine import daterange,known_status,strict_sum
    a,b=interval; wanted=daterange(a,b); covered=set(); notes=[]
    for c in ds.coverage:
        if c['grain']=='items' and (rid in c['restaurant_ids'] or c['all_partners']): covered |= daterange(c['start'],c['end']) & wanted
    observed=[r for r in ds.items if r['restaurant_id']==rid and a<=r['date']<=b]
    parent={r['order_id']:r for r in ds.rows if r['grain']=='orders'}
    accepted=[]; unmatched=0; unnamed=0
    for item in observed:
        order=parent.get(item['order_id'])
        if order:
            valid=order['restaurant_id']==rid and order['date']==item['date']
            status=known_status(order,ds.settings)
            if item.get('status') and known_status(item,ds.settings)!=status: valid=False
            if not valid: unmatched+=1;continue
        else: status=known_status(item,ds.settings)
        if status is None: unmatched+=1;continue
        if status=='delivered':
            accepted.append(item)
            if not item.get('product_id'): unnamed+=1
    exact=covered==wanted and unmatched==0
    if not observed: notes.append('Aucune ligne produit observée sur cette période.')
    if covered!=wanted: notes.append(f'Articles : couverture non certifiée {len(covered)}/{len(wanted)} jours. Absence de produit ≠ zéro vente.')
    if unmatched: notes.append(f'{unmatched} lignes avec statut inconnu ou incohérent avec la commande : exclues, variations N/A.')
    if unnamed: notes.append('Sans Product ID : correspondance exacte nom + catégorie uniquement. Un renommage peut créer deux produits.')
    missing_qty=sum(r.get('quantity') is None for r in accepted)
    if missing_qty: notes.append(f'{missing_qty} lignes sans quantité : classement exhaustif et mix indisponibles.')
    by_product=defaultdict(list);by_category=defaultdict(list)
    for r in accepted:
        by_product[r['product_key']].append(r);by_category[r.get('category') or 'Non catégorisé'].append(r)
    money=bool(ds.settings.get('line_gmv_definition'))
    def total(rr):
        return {'quantity':strict_sum(rr,'quantity'),'gmv':strict_sum(rr,'line_gmv') if money else None}
    products={key:{'name':rr[-1].get('product_name') or key,'category':rr[-1].get('category') or 'Non catégorisé',**total(rr)} for key,rr in by_product.items()}
    cats={key:{'name':key,**total(rr)} for key,rr in by_category.items()}
    if accepted: totals=total(accepted)
    else: totals={'quantity':0 if exact else None,'gmv':0 if exact and money else None}
    return {'products':products,'categories':cats,'totals':totals,'days':len(wanted),'covered_days':len(covered),'exact':exact,'notes':notes,
            'lines':len(accepted),'rankable':missing_qty==0,'sources':sorted({r['source'] for r in observed})}

def compare_products(ds,rid,current,previous):
    from engine import ratio,fmt
    cur,pre=item_period(ds,rid,current),item_period(ds,rid,previous)
    reliable=cur['exact'] and pre['exact']
    def rows(kind):
        result=[]
        for key in set(cur[kind])|set(pre[kind]):
            ca,pa=cur[kind].get(key),pre[kind].get(key)
            meta=ca or pa
            a=ca['quantity'] if ca else (0 if cur['exact'] else None)
            b=pa['quantity'] if pa else (0 if pre['exact'] else None)
            ga=ca['gmv'] if ca else (0 if cur['exact'] and ds.settings.get('line_gmv_definition') else None)
            gb=pa['gmv'] if pa else (0 if pre['exact'] and ds.settings.get('line_gmv_definition') else None)
            perday=a/cur['days'] if a is not None and cur['exact'] else None
            prevday=b/pre['days'] if b is not None and pre['exact'] else None
            delta=perday-prevday if reliable and perday is not None and prevday is not None else None
            rel=delta/prevday if delta is not None and prevday else None
            growth=f'{rel*100:+.1f} %'.replace('.',',') if rel is not None else ('Nouvelle activité' if delta is not None and prevday==0 and perday>0 else 'N/A')
            share=ratio(a,cur['totals']['quantity']);previous_share=ratio(b,pre['totals']['quantity'])
            result.append({'id':key,'name':meta['name'],'category':meta.get('category',''), 'current_qty':a,'previous_qty':b,
                           'current_gmv':ga,'previous_gmv':gb,'qty_per_day':perday,'previous_per_day':prevday,
                           'delta_per_day':delta,'relative':rel,'growth':growth,'share':share,'previous_share':previous_share,
                           'share_delta':share-previous_share if reliable and share is not None and previous_share is not None else None})
        result.sort(key=lambda r: (-(r['current_qty'] if r['current_qty'] is not None else -1),r['id']))
        return result
    products=rows('products');categories=rows('categories')
    top=[r for r in products if r['current_qty'] is not None and r['current_qty']>0][:3] if cur['rankable'] else []
    drops=sorted((r for r in products if r['delta_per_day'] is not None and r['delta_per_day']<0),key=lambda r:(r['delta_per_day'],r['id']))[:3]
    risers=sorted((r for r in products if r['delta_per_day'] is not None and r['delta_per_day']>0),key=lambda r:(-r['delta_per_day'],r['id']))[:3]
    return {'current':cur,'previous':pre,'products':products,'categories':categories,'top3':top,'drops':drops,'risers':risers,'reliable':reliable,
            'top3_share':ratio(sum(r['current_qty'] for r in top),cur['totals']['quantity']),
            'ranking':'Quantités vendues sur commandes livrées. Variations comparées en unités/jour.',
            'definition':ds.settings.get('line_gmv_definition') or 'Montant ligne non défini : GMV produit N/D.'}
