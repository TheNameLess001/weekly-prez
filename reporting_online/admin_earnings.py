"""Admin Earnings adapter. Source amounts take precedence over derived values."""
from datetime import datetime,timedelta
import re
import pandas as pd

DEFINITIONS={
 'ca_definition':'Admin Earnings : revenus bruts observés = restaurant commission + delivery amount + service charge, sur commandes livrées. Hors qualification comptable HT/TTC.',
 'commission_definition':'Admin Earnings : restaurant commission, sur commandes livrées. À défaut : item total − amount to restaurant.',
 'commission_base_definition':'Admin Earnings : item total, panier brut avant remises, sur commandes livrées.',
 'promo_definition':'Admin Earnings : coupon admin + Discount Amount financé par admin. Promotions alimentaires uniquement, hors remise livraison et financement restaurant.',
 'benefit_definition':'Admin Earnings : final earnings déclaré, sur commandes livrées. À défaut : admin earnings + amount to restaurant. Ce montant source ne constitue pas un bénéfice comptable.',
}
FORMULAS={
 'requests':'Nombre de commandes distinctes après dédoublonnage',
 'delivered':'Nombre de commandes au statut livré',
 'cancelled':'Nombre de commandes annulées ou rejetées',
 'users':'Nombre de clients distincts (customer Phone)',
 'delivered_users':'Nombre de clients distincts sur commandes livrées',
 'aov':'GMV livré / commandes livrées',
 'sr':'Commandes livrées / commandes reçues',
 'frequency':'Commandes livrées / clients livrés uniques',
 'repeat_share':'Clients avec ≥ 2 livraisons dans la période / clients livrés uniques',
 'commission_base':'item total (panier brut)',
 'commission':'restaurant commission ; sinon item total − amount to restaurant',
 'ca':'restaurant commission + delivery amount + service charge',
 'promo':'coupon admin + Discount Amount si Discount Type Provider = admin',
 'benefit':'final earnings ; sinon admin earnings + amount to restaurant',
 'accept_min':'Accepted at − (order day + order time)',
 'prep_min':'picked up time − Accepted at (préparation estimée, attente éventuelle incluse)',
 'ready_min':'Ready by Restaurant − Accepted at (préparation déclarée)',
 'wait_min':'picked up time − Ready by Restaurant (attente après prêt, pas attente du livreur sur place)',
 'travel_min':'(delivered day + delivered at) − picked up time',
 'ctd_min':'delivery time(M) ; sinon livraison − création',
}

def is_admin(columns,norm):
    names={norm(c) for c in columns}
    return {'order_id','order_day','item_total','restaurant_id'}.issubset(names)

def derive(record,r,norm,number):
    raw={norm(k):v for k,v in record.items()}
    for key in ('food_item','food_category'):
        v=raw.get(key)
        if v is not None and not pd.isna(v):r[key]=str(v)
    provenance=r.setdefault('derived',{})
    def n(key):return number(raw.get(key))
    def put(key,value,formula):
        if r.get(key) is None and value is not None:r[key]=value;provenance[key]=formula
    put('commission_base',n('item_total'),FORMULAS['commission_base'])
    if n('item_total') is not None and n('amount_to_restaurant') is not None:
        put('commission',n('item_total')-n('amount_to_restaurant'),FORMULAS['commission'])
    components=[r.get('commission'),n('delivery_amount'),n('service_charge')]
    if all(v is not None for v in components):put('ca',sum(components),FORMULAS['ca'])
    if n('admin_earnings') is not None and n('amount_to_restaurant') is not None:
        put('benefit',n('admin_earnings')+n('amount_to_restaurant'),FORMULAS['benefit'])
    coupon,discount=n('coupon_admin'),n('discount_amount')
    provider=norm(raw.get('discount_type_provider',''))
    if coupon is not None and discount is not None and (discount==0 or provider in ('admin','restaurant')):
        put('promo',coupon+(discount if provider=='admin' else 0),FORMULAS['promo'])
    status=norm(r.get('status',''))
    if status in ('restaurant_rejected','cancelled_by_restaurant'):r['cancel_actor']='restaurant'
    elif status in ('cancelled_by_user','cancelled_by_customer'):r['cancel_actor']='client'
    elif status=='cancelled_by_driver':r['cancel_actor']='driver'
    elif status=='cancelled_by_admin':r['cancel_actor']='other'
    # Explicit cancellation actors override this status fallback only if recognized.
    actor=norm(r.get('cancel_actor',''))
    if actor=='customer_app':r['cancel_actor']='client'
    elif actor not in ('restaurant','resto','merchant','client','customer','user','driver','livreur','courier','other','autre'):
        r['cancel_actor']=None
    def stamp(value,base=None):
        if value is None or pd.isna(value) or not str(value).strip():return None
        s=str(value).strip().replace(' / ',' ')
        try:
            if re.fullmatch(r'\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?',s):
                if base is None:return None
                t=pd.to_datetime(s).time();dt=datetime.combine(base.date(),t)
                return dt+timedelta(days=1) if dt<base else dt
            return pd.to_datetime(s,dayfirst=not bool(re.match(r'^\d{4}-',s))).to_pydatetime().replace(tzinfo=None)
        except (ValueError,TypeError,OverflowError):return None
    def dated(day_value,time_value):
        if day_value is None or pd.isna(day_value):return None
        try:
            if isinstance(day_value,(int,float)):
                base=(pd.Timestamp('1899-12-30')+pd.to_timedelta(float(day_value),unit='D')).date()
            else:base=pd.to_datetime(day_value,dayfirst=not bool(re.match(r'^\d{4}-',str(day_value)))).date()
        except (ValueError,TypeError):return None
        return stamp(str(base)+' '+str(time_value))
    order=dated(raw.get('order_day'),raw.get('order_time'))
    accepted=stamp(raw.get('accepted_at'),order)
    ready=stamp(raw.get('ready_by_restaurant'),accepted)
    pickup=stamp(raw.get('picked_up_time'),accepted or order)
    delivered=dated(raw.get('delivered_day'),raw.get('delivered_at'))
    # Reject contradictory event chains. Full dated timestamps are never wrapped.
    accepted_valid=accepted is not None and (order is None or accepted>=order)
    pickup_valid=pickup is not None and (accepted is None or (accepted_valid and pickup>=accepted)) and (delivered is None or pickup<=delivered)
    ready_valid=ready is not None and accepted_valid and ready>=accepted and (pickup is None or ready<=pickup)
    for key,start,end in [('accept_min',order,accepted if accepted_valid else None),
                          ('prep_min',accepted,pickup if pickup_valid else None),
                          ('ready_min',accepted,ready if ready_valid else None),
                          ('wait_min',ready if ready_valid else None,pickup if pickup_valid else None),
                          ('travel_min',pickup if pickup_valid else None,delivered),('ctd_min',order,delivered)]:
        if start and end:
            minutes=(end-start).total_seconds()/60
            if 0<=minutes<=1440:put(key,minutes,FORMULAS[key])
    if not provenance:r.pop('derived',None)


def extract_items(order,ds):
    """Admin export encodes explicit quantities as [Product{2}]. No basket allocation."""
    # Mapped empty CSV/Excel cells can exist as None, NaN or pandas.NA.
    # dict.get(default) does not replace a present missing value.
    def optional_text(value):
        return '' if value is None or pd.isna(value) else str(value).strip()
    text=optional_text(order.get('food_item'))
    pattern=r'\[([^\[\]]+?)\{(\d+(?:[.,]\d+)?)\}\]'
    matches=list(re.finditer(pattern,text))
    remainder=re.sub(pattern,'',text)
    if not matches or remainder.strip(' ,'):
        ds.quality.append({'type':'articles non interprétables','source':order['source'],'detail':f"L{order['line']}: format Food Item non reconnu, aucune quantité inventée"})
        return []
    cats=[c.strip() for c in optional_text(order.get('food_category')).strip('/').split('/') if c.strip()]
    # Export categories can be deduplicated, so never assume positional alignment.
    category=cats[0] if len(cats)==1 else 'Catégorie non attribuable'
    result=[]
    from engine import norm
    for i,m in enumerate(matches):
        name=m.group(1).strip();qty=float(m.group(2).replace(',','.'))
        if not name or qty<=0:return []
        result.append({**{k:order.get(k) for k in ('restaurant_id','order_id','date','status','source','line')},
            'grain':'items','complete':False,'line_id':f"ADMIN:{i}", 'product_id':'ADMIN:'+norm(name),
            'product_name':name,'category':category,'quantity':qty,'line_gmv':None,'from_admin':True})
    return result
