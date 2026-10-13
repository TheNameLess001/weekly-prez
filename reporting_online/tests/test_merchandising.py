import sys,tempfile,unittest
from pathlib import Path
from datetime import date
import pandas as pd
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from engine import inspect_file,ingest,compare,aggregate
from merchandising import compare_products

class ProductTests(unittest.TestCase):
    def setUp(self):self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()
    def row(self,oid,line,pid='P1',qty=1,day='2026-09-21',amount=10,**kw):
        return dict(restaurant_id='A',order_id=oid,line_id=line,product_id=pid,product_name=pid,category='Plats',quantity=qty,date=day,line_gmv=amount,status='delivered',**kw)
    def source(self,rows,name='items.csv',complete=True):
        p=self.root/name;pd.DataFrame(rows).to_csv(p,index=False);sh=inspect_file(p)
        for s in sh:s.update(complete=complete,coverage_start='2026-09-21',coverage_end='2026-10-04',all_partners=True)
        return {'path':str(p),'sheets':sh}
    def result(self,rows,complete=True):
        ds=ingest([self.source(rows,complete=complete)],{'line_gmv_definition':'Montant article TTC livré'})
        return ds,compare_products(ds,'A',(date(2026,9,28),date(2026,10,4)),(date(2026,9,21),date(2026,9,27)))
    def test_line_dedup_and_multiple_items(self):
        one=self.row('O1','L1');two=self.row('O1','L2',qty=2)
        ds,r=self.result([one,one,two]);self.assertEqual(len(ds.items),2);self.assertEqual(r['previous']['totals']['quantity'],3)
        self.assertEqual(len(ds.rows),0);self.assertIsNone(aggregate(ds,'A',(date(2026,9,21),date(2026,9,27)))['values']['requests'])
    def test_category_mix_and_top3(self):
        rows=[self.row('O1','L1','P1',6,'2026-09-28'),self.row('O2','L1','P2',3,'2026-09-28'),self.row('O3','L1','P3',1,'2026-09-28')]
        ds,r=self.result(rows);self.assertEqual([p['id'] for p in r['top3']],['P1','P2','P3']);self.assertEqual(r['top3_share'],1)
        self.assertEqual(r['categories'][0]['share'],1)
    def test_drop_disappeared_product_only_if_exhaustive(self):
        ds,r=self.result([self.row('O1','L1',qty=7)])
        self.assertEqual(r['drops'][0]['growth'],'-100,0 %');self.assertEqual(r['drops'][0]['current_qty'],0)
        ds,r=self.result([self.row('O1','L1',qty=7)],False);self.assertEqual(r['drops'],[])
        self.assertIsNone(r['products'][0]['current_qty'])
    def test_normalized_different_lengths(self):
        rows=[self.row('O1','L1',qty=14),self.row('O2','L1',qty=7,day='2026-09-28')]
        ds,r=self.result(rows)
        v=compare_products(ds,'A',(date(2026,9,28),date(2026,9,30)),(date(2026,9,21),date(2026,9,27)))['products'][0]
        self.assertAlmostEqual(v['delta_per_day'],7/3-2);self.assertGreater(v['relative'],0)
    def test_missing_quantity_no_fake_top(self):
        ds,r=self.result([self.row('O1','L1',qty=None,day='2026-09-28')]);self.assertEqual(r['top3'],[]);self.assertIsNone(r['current']['totals']['quantity'])
    def test_conflicts_invalidate_item_coverage(self):
        ds,r=self.result([self.row('O1','L1'),self.row('O1','L1',qty=3)])
        self.assertEqual(ds.items,[]);self.assertFalse(r['reliable'])
    def test_parent_order_status_prevents_cancelled_sale(self):
        item=self.source([self.row('O1','L1')]);p=self.root/'orders.csv';pd.DataFrame([dict(restaurant_id='A',order_id='O1',date='2026-09-21',status='cancelled')]).to_csv(p,index=False)
        ds=ingest([item,{'path':str(p),'sheets':inspect_file(p)}]);r=compare_products(ds,'A',(date(2026,9,21),date(2026,9,27)),(date(2026,9,14),date(2026,9,20)))
        self.assertEqual(r['current']['lines'],0);self.assertFalse(r['current']['exact'])
    def test_unknown_status_not_zero(self):
        row=self.row('O1','L1');row.pop('status');ds,r=self.result([row]);self.assertFalse(r['previous']['exact']);self.assertIsNone(r['previous']['totals']['quantity'])

if __name__=='__main__':unittest.main(verbosity=2)
