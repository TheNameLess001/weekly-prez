"""Regression tests for observed PoP, timing reconstruction and explicit item quantities."""
import sys,tempfile,unittest
from pathlib import Path
from datetime import date
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from engine import ingest,inspect_file,compare,aggregate,norm,number
from admin_earnings import derive

class ComparisonTests(unittest.TestCase):
    def dataset(self,rows,settings=None):
        tmp=tempfile.TemporaryDirectory();self.addCleanup(tmp.cleanup)
        p=Path(tmp.name)/'orders.csv';pd.DataFrame(rows).to_csv(p,index=False)
        return ingest([{'path':str(p),'sheets':inspect_file(p)}],settings)
    def row(self,oid,day='2026-10-05',**kw):
        return dict(restaurant_id='A',restaurant_name='Test',order_id=oid,date=day,status='delivered',user_id='U'+oid,gmv=100,**kw)
    def report(self,ds):return compare(ds,'A',(date(2026,10,5),date(2026,10,11)),(date(2026,9,28),date(2026,10,4)))
    def test_uncertified_both_periods_compare_and_normalize(self):
        ds=self.dataset([self.row('1','2026-09-28'),self.row('2'),self.row('3')])
        r=self.report(ds);m={m['key']:m for m in r['metrics']}
        self.assertEqual(r['comparison_basis'],'observée');self.assertEqual(m['delivered']['growth'],'+100,0 %')
        self.assertEqual(m['gmv']['delta'],100);self.assertEqual(m['active_days']['current'],1)
        self.assertAlmostEqual(m['delivered_per_day']['current'],2/7)
        short=compare(ds,'A',(date(2026,10,5),date(2026,10,5)),(date(2026,9,28),date(2026,10,4)))
        self.assertAlmostEqual(next(m for m in short['metrics'] if m['key']=='delivered_per_day')['relative'],13)
    def test_rate_points_and_relative_percent_are_distinct(self):
        rows=[self.row(str(i),'2026-09-28',**({'cancel_actor':'client'} if i==0 else {})) for i in range(10)]
        rows[0]['status']='cancelled';rows[1]['status']='cancelled';rows[1]['cancel_actor']='client'
        rows += [self.row('C'+str(i)) for i in range(10)];rows[-1]['status']='cancelled';rows[-1]['cancel_actor']='client'
        r=self.report(self.dataset(rows));m={x['key']:x for x in r['metrics']}
        self.assertEqual(m['sr']['delta_text'],'+10,0 pts');self.assertEqual(m['sr']['growth'],'+12,5 %')
        self.assertEqual(m['cancel_rate']['growth'],'-50,0 %');self.assertTrue(m['cancel_rate']['favorable'])
    def test_negative_reference_no_reversed_percent(self):
        ds=self.dataset([self.row('1','2026-09-28',benefit=-20),self.row('2',benefit=-10)],{'benefit_definition':'Source signed'})
        m=next(x for x in self.report(ds)['metrics'] if x['key']=='benefit')
        self.assertEqual(m['delta'],10);self.assertEqual(m['growth'],'N/A (base négative)');self.assertTrue(m['favorable'])
    def test_unknown_actor_is_reconciled(self):
        a=self.row('1',cancel_actor='restaurant');a['status']='cancelled'
        b=self.row('2',cancel_actor='unrecognized');b['status']='cancelled'
        v=aggregate(self.dataset([a,b]),'A',(date(2026,10,5),date(2026,10,5)))['values']
        self.assertEqual(v['cancelled'],2);self.assertEqual(v['cancel_restaurant'],1);self.assertEqual(v['cancel_unknown'],1)
        self.assertEqual(sum(v['cancel_'+a] for a in ('restaurant','driver','client','other','unknown')),v['cancelled'])
    def test_observed_new_customers_distinct_from_global_new(self):
        rows=[self.row('1','2026-09-28'),self.row('2')];rows[1]['user_id']='U1';rows.append(self.row('3'))
        v=self.report(self.dataset(rows))['current']['values']
        self.assertEqual(v['observed_new_users'],1);self.assertIsNone(v['new_users']);self.assertEqual(v['delivered_users'],2)
    def test_sla_only_with_explicit_threshold(self):
        rows=[self.row('1',ctd_min=20),self.row('2',ctd_min=50)]
        self.assertIsNone(self.report(self.dataset(rows))['current']['values']['sla_rate'])
        r=self.report(self.dataset(rows,{'sla_target_min':40}))
        self.assertEqual(r['current']['values']['sla_rate'],.5);self.assertEqual(r['current']['samples']['sla_rate'],2)
    def test_invalid_duration_does_not_discard_order(self):
        r=self.report(self.dataset([self.row('1',ctd_min=-2)]))
        self.assertEqual(r['current']['values']['gmv'],100);self.assertIsNone(r['current']['values']['ctd_min'])
    def test_reconstruct_times_with_missing_ready_and_source_ctd(self):
        raw={'order day':'2026-10-05','order time':'23:55:00','Accepted at':'2026-10-05 / 23:57:00','picked up time':'00:08:00','delivered day':'2026-10-06','delivered at':'00:18:00'}
        r={};derive(raw,r,norm,number)
        self.assertEqual(r['prep_min'],11);self.assertEqual(r['ctd_min'],23);self.assertEqual(r['travel_min'],10)
        self.assertNotIn('wait_min',r);self.assertNotIn('ready_min',r)
        raw['Ready by Restaurant']='2026-10-05 / 23:00:00';r={'ctd_min':24};derive(raw,r,norm,number)
        self.assertEqual(r['ctd_min'],24);self.assertEqual(r['prep_min'],11);self.assertNotIn('ready_min',r)
    def test_food_item_quantities_and_observed_product_comparison(self):
        raw={'Restaurant ID':'A','restaurant name':'Test','order id':'1','order day':'2026-09-28','status':'Delivered','item total':100,'Food Item':'[Burger{2}],[Pizza{1}]','Food Category':'Burgers/Pizzas/'}
        newer={**raw,'order id':'2','order day':'2026-10-05','Food Item':'[Burger{1}],[Pizza{3}]'}
        ds=self.dataset([raw,newer]);r=self.report(ds);m=r['merchandising']
        self.assertEqual(m['current']['totals']['quantity'],4);self.assertEqual(m['previous']['totals']['quantity'],3)
        self.assertEqual(m['top3'][0]['name'],'Pizza');self.assertEqual(m['drops'][0]['name'],'Burger');self.assertEqual(m['drops'][0]['growth'],'-50,0 %')
        self.assertEqual(m['risers'][0]['growth'],'+200,0 %');self.assertIsNone(m['current']['totals']['gmv'])
        self.assertEqual(len(ds.rows),2);self.assertEqual(r['current']['values']['gmv'],100)
        self.assertEqual(m['products'][0]['category'],'Catégorie non attribuable')
    def test_malformed_food_is_not_invented(self):
        raw={'Restaurant ID':'A','order id':'1','order day':'2026-10-05','status':'Delivered','item total':100,'Food Item':'Burger, Pizza'}
        ds=self.dataset([raw]);self.assertEqual(ds.items,[]);self.assertEqual(len(ds.rows),1)

    def test_explicit_items_prevent_auto_quantity_double_count(self):
        raw={'Restaurant ID':'A','order id':'1','order day':'2026-10-05','status':'Delivered','item total':100,'Food Item':'[Burger{2}]'}
        ds=self.dataset([raw])
        from engine import ingest,inspect_file
        tmp=tempfile.TemporaryDirectory();self.addCleanup(tmp.cleanup);root=Path(tmp.name)
        orders=root/'orders.csv';items=root/'items.csv';pd.DataFrame([raw]).to_csv(orders,index=False)
        pd.DataFrame([dict(restaurant_id='A',order_id='1',date='2026-10-05',line_id='L1',product_id='B',product_name='Burger',quantity=2,status='delivered')]).to_csv(items,index=False)
        ds=ingest([{'path':str(p),'sheets':inspect_file(p)} for p in (orders,items)])
        self.assertEqual(len(ds.items),1);self.assertEqual(self.report(ds)['merchandising']['current']['totals']['quantity'],2)

    def test_french_timestamps_do_not_swap_month_and_day(self):
        raw={'order day':'04/10/2026','order time':'10:00:00','Accepted at':'04/10/2026 10:05:00','picked up time':'10:25:00','delivered day':'04/10/2026','delivered at':'10:35:00'}
        r={};derive(raw,r,norm,number);self.assertEqual(r['accept_min'],5);self.assertEqual(r['prep_min'],20);self.assertEqual(r['ctd_min'],35)
