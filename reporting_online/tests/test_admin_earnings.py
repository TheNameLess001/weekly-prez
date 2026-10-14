import sys,unittest,tempfile
from pathlib import Path
from datetime import date
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import pandas as pd
from engine import inspect_file,ingest,aggregate,norm,number
from admin_earnings import derive

class AdminTests(unittest.TestCase):
    def fixture(self):
        return {'Restaurant ID':'TEST','restaurant name':'Restaurant test','order id':'O1','customer Phone':'TEST_U1','order day':'2026-10-05','order time':'23:55:00','Accepted at':'2026-10-05 / 23:57:00','Ready by Restaurant':'2026-10-06 / 00:05:00','picked up time':'00:08:00','delivered day':'2026-10-06','delivered at':'00:18:00','status':'Delivered','item total':100,'restaurant commission':20,'amount to restaurant':80,'delivery amount':8,'service charge':3,'coupon admin':2,'Discount Amount':10,'Discount Type Provider':'admin','admin earnings':-70,'final earnings':10,'delivery time(M)':23}
    def test_auto_mapping_formulas_and_statuses(self):
        with tempfile.TemporaryDirectory() as t:
            p=Path(t)/'admin.csv';a=self.fixture();b={**a,'order id':'O2','status':'Cancelled by user'};c={**a,'order id':'O3','status':'Restaurant Rejected'}
            pd.DataFrame([a,b,c]).to_csv(p,index=False)
            sheets=inspect_file(p);self.assertEqual(sheets[0]['grain'],'orders');self.assertEqual(sheets[0]['profile'],'admin_earnings');self.assertEqual(sheets[0]['mapping']['date'],'order day')
            ds=ingest([{'path':str(p),'sheets':sheets}]);self.assertEqual(len(ds.rows),3)
            v=aggregate(ds,'TEST',(date(2026,10,5),date(2026,10,5)))['values']
            for k,n in {'requests':3,'delivered':1,'cancelled':2,'cancel_client':1,'cancel_restaurant':1,'gmv':100,'commission':20,'ca':31,'promo':12,'benefit':10,'users':1,'aov':100,'frequency':1,'accept_min':2,'prep_min':11,'ready_min':8,'wait_min':3,'travel_min':10}.items():self.assertAlmostEqual(v[k],n,msg=k)
            self.assertIsNone(v['new_users'])
    def test_source_precedence_and_missing_inputs(self):
        raw=self.fixture();raw['final earnings']=None;raw['restaurant commission']=None
        r={};derive(raw,r,norm,number)
        self.assertEqual(r['commission'],20);self.assertEqual(r['benefit'],10)
        r={'benefit':99,'commission':25};derive(raw,r,norm,number);self.assertEqual(r['benefit'],99);self.assertEqual(r['commission'],25)
        raw['admin earnings']=None;raw['delivery amount']=None;raw['coupon admin']=None
        r={};derive(raw,r,norm,number)
        self.assertNotIn('benefit',r);self.assertNotIn('ca',r);self.assertNotIn('promo',r)
    def test_restaurant_promotions_not_admin_cost(self):
        raw=self.fixture();raw['Discount Type Provider']='restaurant';r={};derive(raw,r,norm,number);self.assertEqual(r['promo'],2)
        raw['Discount Type Provider']='unknown';r={};derive(raw,r,norm,number);self.assertNotIn('promo',r)
    def test_invalid_times_not_wrapped_and_excel_dates(self):
        raw=self.fixture();raw['order day']=pd.Timestamp('2026-10-05');raw['Accepted at']='2026-10-05 / 23:50:00';r={};derive(raw,r,norm,number);self.assertNotIn('accept_min',r)
        raw=self.fixture();raw['order day']=pd.Timestamp('2026-10-05');r={};derive(raw,r,norm,number);self.assertEqual(r['accept_min'],2)
