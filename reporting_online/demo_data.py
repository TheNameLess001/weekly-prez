"""Reproducible fictitious observations. Not a Yassir performance source."""
import csv
import json
import random
from datetime import date,timedelta
from pathlib import Path

DEFINITIONS={
 'ca_definition':'DÉMO : commission brute, même montant que le champ commission. Aucun revenu restaurant.',
 'commission_definition':'DÉMO : 20 % du GMV TTC des commandes livrées.',
 'commission_base_definition':'DÉMO : GMV livré TTC.',
 'promo_definition':'DÉMO : coût promotionnel Yassir sur commandes livrées.',
 'benefit_definition':'DÉMO : commission moins promotion et coût livraison fictif de 6 MAD. Champ contribution fourni.',
 'tax_basis':'DÉMO : GMV et base TTC. Commission brute illustrative, aucune comptabilité réelle.',
 'verified_ids':{},
 'line_gmv_definition':'DÉMO : montant TTC réparti sur les lignes produits livrées, hors autre revenu. Articles et catégories fictifs.',
}

def generate(folder):
    folder=Path(folder); folder.mkdir(parents=True,exist_ok=True); rng=random.Random(202640)
    rows=[]; registries=[{'restaurant_id':'DEMO_BELHAJ','restaurant_name':'Belhaj','city':'Casablanca','brand':'Belhaj','is_demo':1},
                         {'restaurant_id':'DEMO_RABAT','restaurant_name':'Partenaire Démo Rabat','city':'Rabat','brand':'Enseigne Démo','is_demo':1},
                         {'restaurant_id':'DEMO_ZERO','restaurant_name':'Partenaire sans commande (démo)','city':'Casablanca','brand':'Enseigne Démo','is_demo':1}]
    for meta in registries[:2]:
        for di in range(14):
            dt=date(2026,9,21)+timedelta(days=di); current=di>=7
            n=(23 if current else 18)+(di%4) if meta['restaurant_id']=='DEMO_BELHAJ' else 7+di%3
            for j in range(n):
                delivered=rng.random()<(0.85 if current else 0.78)
                uid='U'+str(rng.randint(1,95 if current else 82)); gmv=round(rng.uniform(55,190),2)
                is_new=int(int(uid[1:])>82); first='2026-09-28' if is_new else '2026-01-05'
                prep=round(rng.uniform(9,20) if current else rng.uniform(12,27),1)
                accept=round(rng.uniform(1,4),1); wait=round(rng.uniform(2,8),1); travel=round(rng.uniform(8,22),1); ctd=round(accept+prep+wait+travel,1)
                commission=round(gmv*.2,2); promo=5. if current else 7.
                rows.append({**meta,'order_id':f"{meta['restaurant_id']}_{di}_{j}",'date':str(dt),'status':'delivered' if delivered else 'cancelled',
                             'user_id':uid,'user_first_date':first,'is_new':is_new,
                             'gmv':gmv if delivered else '', 'commission':commission if delivered else '', 'ca':commission if delivered else '',
                             'commission_base':gmv if delivered else '', 'promo':promo if delivered else '',
                             'benefit':round(commission-promo-6,2) if delivered else '',
                             'cancel_actor':'' if delivered else rng.choice(['restaurant','driver','client']),
                             'cancel_reason':'' if delivered else rng.choice(['Rupture produit','Restaurant fermé','Client indisponible']),
                             'accept_min':accept if delivered else '', 'prep_min':prep if delivered else '', 'wait_min':wait if delivered else '',
                             'travel_min':travel if delivered else '', 'ctd_min':ctd if delivered else '', 'sla_met':int(ctd<=45) if delivered else ''})
    # Copy 5 observations to exercise safe order-ID deduplication.
    item_rng=random.Random(202641);items=[]
    menu=[('P1','Tajine poulet','Plats'),('P2','Couscous légumes','Plats'),('P3','Tajine kefta','Plats'),
          ('P4','Sandwich poulet','Sandwichs'),('P5','Salade marocaine','Accompagnements'),('P6','Thé à la menthe','Boissons'),
          ('P7','Jus d’orange','Boissons'),('P8','Dessert maison','Desserts')]
    for order in rows:
        if order['status']!='delivered':continue
        current=order['date']>='2026-09-28'
        basket=item_rng.choices(menu,weights=[8,2 if current else 14,7,13 if current else 2,3,2 if current else 7,5,4],k=item_rng.randint(1,3))
        quantities=[item_rng.choice([1,1,2]) for _ in basket];cents=round(order['gmv']*100)
        allocated=[int(cents*q/sum(quantities)) for q in quantities];allocated[-1]+=cents-sum(allocated)
        for li,((pid,name,cat),q,amount) in enumerate(zip(basket,quantities,allocated)):
            items.append({**{k:order[k] for k in ('restaurant_id','restaurant_name','city','brand','is_demo','order_id','date','status')},
                          'line_id':order['order_id']+f'_L{li+1}','product_id':pid,'product_name':name,'category':cat,'quantity':q,'line_gmv':amount/100})
    rows+=rows[:5]
    items+=items[:3]
    for name,items in [('orders_DEMO.csv',rows),('registry_DEMO.csv',registries),('items_DEMO.csv',items)]:
        with (folder/name).open('w',newline='',encoding='utf-8-sig') as f:
            w=csv.DictWriter(f,fieldnames=list(items[0])); w.writeheader(); w.writerows(items)
    (folder/'definitions.json').write_text(json.dumps(DEFINITIONS,ensure_ascii=False,indent=2),encoding='utf-8')
    (folder/'LISEZ_MOI.txt').write_text('DONNÉES ENTIÈREMENT FICTIVES. Générées pour tester le logiciel. Aucun résultat réel de Belhaj ou Yassir. Dates 21/09/2026 au 04/10/2026.\n',encoding='utf-8')

if __name__=='__main__': generate(Path(__file__).parent/'demo')
