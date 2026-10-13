import json
import socket
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from engine import inspect_file,ingest,aggregate,compare,periods,number,day
from deck import export_batch,export_one

class EngineTests(unittest.TestCase):
    def setUp(self): self.tmp=tempfile.TemporaryDirectory(); self.root=Path(self.tmp.name)
    def tearDown(self): self.tmp.cleanup()
    def src(self,name,rows,grain='orders',start='2026-09-21',end='2026-09-27',complete=True):
        p=self.root/name; df=pd.DataFrame(rows)
        if p.suffix=='.xlsx': df.to_excel(p,index=False)
        else: df.to_csv(p,index=False)
        s=inspect_file(p)
        for sh in s: sh.update(grain=grain,complete=complete,coverage_start=start,coverage_end=end,all_partners=True)
        return {'path':str(p),'sheets':s}
    def order(self,oid='1',**extra):
        return {'restaurant_id':'A','restaurant_name':'Belhaj','order_id':oid,'date':'2026-09-21','status':'delivered','user_id':'U1',
                'gmv':100,'ctd_min':30,'commission':20,**extra}
    def week(self,**extra):
        return {'restaurant_id':'A','restaurant_name':'Belhaj','start':'2026-09-21','end':'2026-09-27','requests':10,'delivered':8,'gmv':800,**extra}
    def values(self,ds,a='2026-09-21',b='2026-09-27'): return aggregate(ds,'A',(day(a),day(b)))['values']
    def test_duplicates_and_users(self):
        s=self.src('orders.csv',[self.order(),self.order(),self.order('2',gmv=200),self.order('3',status='cancelled',gmv=None)])
        ds=ingest([s]); v=self.values(ds)
        self.assertEqual(v['requests'],3); self.assertEqual(v['delivered'],2); self.assertEqual(v['users'],1)
        self.assertEqual(v['gmv'],300); self.assertEqual(v['aov'],150); self.assertAlmostEqual(v['sr'],2/3)
        self.assertEqual(v['repeat_share'],1); self.assertEqual(v['frequency'],2)
        self.assertEqual(sum(q['type']=='doublon exclu' for q in ds.quality),1)
    def test_conflicting_order_quarantined(self):
        ds=ingest([self.src('orders.csv',[self.order(),self.order(gmv=200)])])
        self.assertEqual(len(ds.rows),0); self.assertIsNone(self.values(ds)['requests'])
    def test_order_and_summary_never_added(self):
        ds=ingest([self.src('orders.csv',[self.order()]),self.src('summary.csv',[self.week()],grain='weekly')])
        self.assertEqual(self.values(ds)['requests'],1)
        self.assertTrue(any(q['type']=='récapitulatif exclu' for q in ds.quality))
    def test_missing_is_not_zero(self):
        ds=ingest([self.src('orders.csv',[self.order(gmv=None)])]); self.assertIsNone(self.values(ds)['gmv'])
        self.assertIsNone(self.values(ds)['ca']); self.assertIsNone(self.values(ds)['commission'])
    def test_negative_contribution_valid(self):
        ds=ingest([self.src('orders.csv',[self.order(benefit=-7)])],{'benefit_definition':'Commission - coûts'})
        self.assertEqual(self.values(ds)['benefit'],-7)
    def test_demo_marker_survives_reimport(self):
        ds=ingest([self.src('uploaded.csv',[self.order(is_demo=1)])])
        self.assertTrue(ds.demo)
    def test_partial_week_refuses_prorata(self):
        ds=ingest([self.src('weekly.csv',[self.week()],grain='weekly')])
        result=aggregate(ds,'A',(date(2026,9,22),date(2026,9,24)))
        self.assertIsNone(result['values']['gmv']); self.assertIn('aucun découpage',result['notes'][0])
    def test_weighted_mean_and_median(self):
        rows=[self.week(ctd_min=20,ctd_min_n=2,ctd_min_median=19),self.week(start='2026-09-28',end='2026-10-04',ctd_min=40,ctd_min_n=8,ctd_min_median=38)]
        ds=ingest([self.src('weekly.xlsx',rows,grain='weekly',end='2026-10-04')]); v=self.values(ds,b='2026-10-04')
        self.assertEqual(v['ctd_min'],36); self.assertIsNone(v['ctd_min_median'])
        self.assertEqual(v['sr'],.8); self.assertEqual(v['aov'],100)
    def test_weekly_unique_not_summed(self):
        rows=[self.week(users=8),self.week(start='2026-09-28',end='2026-10-04',users=6)]
        ds=ingest([self.src('weekly.csv',rows,grain='weekly',end='2026-10-04')]); self.assertIsNone(self.values(ds,b='2026-10-04')['users'])
    def test_zero_reference(self):
        ds=ingest([self.src('orders.csv',[self.order(date='2026-09-28')],start='2026-09-21',end='2026-10-04')])
        report=compare(ds,'A',(date(2026,9,28),date(2026,10,4)),(date(2026,9,21),date(2026,9,27)))
        m=next(m for m in report['metrics'] if m['key']=='gmv')
        self.assertEqual(m['growth'],'Nouvelle activité'); self.assertIsNone(m['relative'])
    def test_uncertified_deltas_disabled(self):
        ds=ingest([self.src('orders.csv',[self.order()],complete=False)])
        report=compare(ds,'A',(date(2026,9,21),date(2026,9,27)),(date(2026,9,14),date(2026,9,20)))
        self.assertTrue(all(m['delta'] is None for m in report['metrics']))
    def test_registry_zero_only_if_exhaustive(self):
        ds=ingest([self.src('orders.csv',[self.order()]),self.src('registry.csv',[{'restaurant_id':'B','restaurant_name':'Zero'}],grain='registry')])
        v=aggregate(ds,'B',(date(2026,9,21),date(2026,9,27)))['values']; self.assertEqual(v['requests'],0); self.assertIsNone(v['aov'])
    def test_name_requires_verified_id(self):
        row=self.order(restaurant_id=None,city='Casa')
        s=self.src('orders.csv',[row]); self.assertEqual(len(ingest([s]).rows),0)
        ds=ingest([s],{'verified_ids':{'belhaj|casa':'A'}}); self.assertEqual(len(ds.rows),1)
    def test_aggregate_overlap_excluded(self):
        rows=[self.week(),self.week(start='2026-09-24',end='2026-09-30')]
        ds=ingest([self.src('weeks.csv',rows,grain='weekly',end='2026-09-30')]); self.assertEqual(len(ds.rows),0)
    def test_financial_complement_not_added(self):
        comp=self.week(requests=None,delivered=None,gmv=None,commission=27)
        ds=ingest([self.src('orders.csv',[self.order(commission=None)]),self.src('finance.csv',[comp],grain='complement')],{'commission_definition':'Commission TTC validée'})
        v=self.values(ds); self.assertEqual(v['requests'],1); self.assertEqual(v['commission'],27)
    def test_financial_complement_conflict_nd(self):
        comp=self.week(requests=None,delivered=None,gmv=None,commission=27)
        ds=ingest([self.src('orders.csv',[self.order()]),self.src('finance.csv',[comp],grain='complement')],{'commission_definition':'Commission validée'})
        self.assertIsNone(self.values(ds)['commission'])
    def test_dates_numbers(self):
        self.assertEqual(number('1 250,35'),1250.35); self.assertEqual(number('1,250.35'),1250.35)
        self.assertEqual(number('1.250,35'),1250.35); self.assertEqual(day('04/10/2026'),date(2026,10,4))
        self.assertEqual(day(46299),date(2026,10,4))
    def test_periods(self):
        c,p=periods('wow','2026-10-05'); self.assertEqual(c,(date(2026,9,28),date(2026,10,4))); self.assertEqual(p,(date(2026,9,21),date(2026,9,27)))
        c,p=periods('wtd','2026-10-01'); self.assertEqual((c[1]-c[0]).days,(p[1]-p[0]).days)
    def test_bulk_isolation_no_network_no_overwrite(self):
        ds=ingest([self.src('orders.csv',[self.order()])])
        interval=(date(2026,9,21),date(2026,9,27))
        with patch.object(socket.socket,'connect',side_effect=AssertionError('Internet interdit')),patch('urllib.request.urlopen',side_effect=AssertionError('Internet interdit')):
            result=export_batch(ds,['missing','A'],interval,interval,self.root,pdf=False)
            self.assertFalse(result['results'][0]['ok']); self.assertTrue(result['results'][1]['ok']); self.assertTrue(Path(result['zip']).is_file())
            report=compare(ds,'A',interval,interval)
            a=export_one(report,self.root); b=export_one(report,self.root)
            self.assertNotEqual(a['pptx'],b['pptx'])
            from pptx import Presentation
            self.assertEqual(len(Presentation(a['pptx']).slides),14)

if __name__=='__main__': unittest.main(verbosity=2)
