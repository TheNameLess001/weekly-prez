import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch
from streamlit.testing.v1 import AppTest
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))

class StreamlitTests(unittest.TestCase):
    def app(self):return AppTest.from_file(str(ROOT/'streamlit_app.py'),default_timeout=40)
    def test_demo_and_full_export_flow(self):
        with patch.dict(os.environ,{'APP_PASSWORD':''}):
            at=self.app().run();self.assertEqual(len(at.exception),0)
            at.button(key='load_demo').click().run();self.assertEqual(len(at.exception),0)
            self.assertEqual(at.session_state['selected_ids'],['DEMO_BELHAJ']);self.assertEqual(len(at.session_state['workspace'].dataset.rows),417)
            self.assertTrue(any('395' in str(table.value) or 'Sandwich' in str(table.value) for table in at.dataframe))
            at.checkbox(key='pdf_export').set_value(False).run()
            at.button(key='generate').click().run(timeout=40);self.assertEqual(len(at.exception),0)
            self.assertEqual(len(at.session_state['workspace'].results),1)
            self.assertEqual(at.session_state['workspace'].results[0]['results'][0]['ok'],True)
    def test_dates_auto_detected_and_bounded(self):
        from datetime import date
        with patch.dict(os.environ,{'APP_PASSWORD':''}):
            at=self.app().run()
            self.assertFalse(any(x.label=='Actuel début' for x in at.date_input))
            at.button(key='load_demo').click().run()
            self.assertEqual(len(at.exception),0)
            self.assertTrue(any('21/09/2026' in x.value and '04/10/2026' in x.value for x in at.success))
            for key in ('current_start','current_end','previous_start','previous_end'):
                widget=at.date_input(key=key)
                self.assertEqual(widget.min,date(2026,9,21));self.assertEqual(widget.max,date(2026,10,4))
            at.selectbox(key='period_mode').set_value('WoW · dernière semaine complète').run()
            self.assertEqual(len(at.exception),0)
            self.assertTrue(any('2026-09-28 au 2026-10-04' in x.value for x in at.caption))
            at.button(key='clear_session').click().run()
            self.assertEqual(len(at.exception),0)
            self.assertFalse(any(x.label=='Périodes disponibles' for x in at.selectbox))
    def test_admin_upload_displays_calculated_fields(self):
        from test_admin_earnings import AdminTests
        import pandas as pd
        with patch.dict(os.environ,{'APP_PASSWORD':'faux-secret-test-import'}):
            at=self.app().run();at.text_input(key='login_password').set_value('faux-secret-test-import');at.button[0].click().run()
            raw=pd.DataFrame([AdminTests().fixture()]).to_csv(index=False).encode()
            at.file_uploader(key='data_upload').set_value([('admin.csv',raw,'text/csv')]).run()
            at.button(key='import_files').click().run()
            self.assertEqual(len(at.exception),0)
            self.assertTrue(any('Admin Earnings reconnu' in x.value for x in at.success))
            requests=next(x for x in at.selectbox if x.label=='requests')
            self.assertEqual(requests.value,'Calculé automatiquement')
            ds=at.session_state['workspace'].dataset
            self.assertEqual(ds.rows[0]['gmv'],100)
            self.assertEqual(ds.rows[0]['ca'],31)
            self.assertEqual(len(ds.items),0)
    def test_targeted_admin_presentations_custom_periods(self):
        from datetime import date
        from test_admin_earnings import AdminTests
        import pandas as pd,json,zipfile
        with patch.dict(os.environ,{'APP_PASSWORD':'faux-secret-test-import'}):
            at=self.app().run();at.text_input(key='login_password').set_value('faux-secret-test-import');at.button[0].click().run()
            rows=[]
            for rid in ('A','B','C'):
                for i,day in enumerate(('2026-09-22','2026-09-24','2026-09-29','2026-10-02')):
                    rows.append({**AdminTests().fixture(),'Restaurant ID':rid,'restaurant name':'Restaurant '+rid,'order id':rid+str(i),'order day':day,'item total':100+i,'amount to restaurant':80+i})
            raw=pd.DataFrame(rows).to_csv(index=False).encode()
            at.file_uploader(key='data_upload').set_value([('admin.csv',raw,'text/csv')]).run();at.button(key='import_files').click().run()
            at.multiselect(key='selected_ids').set_value(['A','C']).run()
            at.selectbox(key='period_mode').set_value('Deux périodes libres').run()
            chosen={'current_start':date(2026,9,29),'current_end':date(2026,10,2),'previous_start':date(2026,9,22),'previous_end':date(2026,9,24)}
            for key,value in chosen.items():at.date_input(key=key).set_value(value)
            at.run();self.assertEqual(len(at.exception),0)
            at.checkbox(key='pdf_export').set_value(False).run();at.button(key='generate').click().run()
            self.assertEqual(len(at.exception),0)
            result=at.session_state['workspace'].results[-1]
            self.assertEqual(result['request']['ids'],['A','C'])
            self.assertEqual(result['request']['current'],(chosen['current_start'],chosen['current_end']))
            self.assertEqual(result['request']['previous'],(chosen['previous_start'],chosen['previous_end']))
            self.assertEqual(len(result['results']),2);self.assertTrue(all(r['ok'] for r in result['results']))
            with zipfile.ZipFile(result['zip']) as z:
                decks=[n for n in z.namelist() if n.endswith('.pptx')];self.assertEqual(len(decks),2)
                for n in decks:
                    import io
                    with zipfile.ZipFile(io.BytesIO(z.read(n))) as deck:
                        self.assertEqual(sum(n.startswith('ppt/slides/slide') and n.endswith('.xml') for n in deck.namelist()),20)
                reports=[json.loads(z.read(n)) for n in z.namelist() if n.endswith('.json') and n not in ('journal.json','qualite.json')]
                self.assertEqual({r['restaurant']['id'] for r in reports},{'A','C'})
                self.assertTrue(all(r['current']['values']['delivered']==2 and r['previous']['values']['delivered']==2 for r in reports))

    def test_legacy_registry_session_upgraded(self):
        from test_admin_earnings import AdminTests
        import pandas as pd
        with patch.dict(os.environ,{'APP_PASSWORD':''}):
            at=self.app().run();ws=at.session_state['workspace']
            raw=pd.DataFrame([AdminTests().fixture()]).to_csv(index=False).encode()
            ws.add_uploads([('legacy.csv',raw)])
            sheet=ws.sources[0]['sheets'][0];sheet.pop('profile',None);sheet['grain']='registry';sheet['mapping'].pop('date');sheet['mapping'].pop('user_id')
            at.session_state['import_schema_release']='old';at.run()
            self.assertEqual(len(at.exception),0)
            self.assertEqual(sheet['grain'],'orders');self.assertEqual(sheet['mapping']['date'],'order day')
            self.assertEqual(len(ws.dataset.rows),1)
    def test_password_gate(self):
        with patch.dict(os.environ,{'APP_PASSWORD':'exemple-test-long-été'}):
            at=self.app().run();self.assertEqual(len(at.exception),0)
            self.assertEqual(len(at.file_uploader),0)
            at.text_input(key='login_password').set_value('incorrect');at.button[0].click().run()
            self.assertTrue(any('incorrect' in e.value for e in at.error))
            at.text_input(key='login_password').set_value('exemple-test-long-été');at.button[0].click().run()
            self.assertEqual(len(at.exception),0);self.assertTrue(at.session_state['authenticated'])
            self.assertGreater(len(at.file_uploader),0)
    def test_two_streamlit_sessions(self):
        with patch.dict(os.environ,{'APP_PASSWORD':''}):
            a=self.app().run();b=self.app().run();a.button(key='load_demo').click().run()
            self.assertNotEqual(a.session_state['workspace'].root,b.session_state['workspace'].root)
            self.assertEqual(b.session_state['workspace'].dataset.rows,[])
    def test_authenticated_file_upload_and_mapping(self):
        from datetime import date
        with patch.dict(os.environ,{'APP_PASSWORD':'faux-secret-test-import'}):
            at=self.app().run();at.text_input(key='login_password').set_value('faux-secret-test-import');at.button[0].click().run()
            raw=b'restaurant_id,restaurant_name,order_id,date,status,user_id,gmv\nTEST,Restaurant test,O1,2026-10-05,delivered,U1,100\n'
            at.file_uploader(key='data_upload').set_value([('test.csv',raw,'text/csv')]).run()
            at.button(key='import_files').click().run();self.assertEqual(len(at.exception),0)
            self.assertEqual(len(at.session_state['workspace'].dataset.rows),1)
            next(x for x in at.checkbox if x.label=='Source exhaustive sur ces dates').set_value(True)
            next(x for x in at.date_input if x.label=='Début couvert').set_value(date(2026,10,5))
            next(x for x in at.date_input if x.label=='Fin couverte').set_value(date(2026,10,11))
            next(x for x in at.button if x.label=='Appliquer le mapping et les dates').click().run()
            self.assertEqual(len(at.exception),0);self.assertEqual(len(at.session_state['workspace'].dataset.coverage),1)

if __name__=='__main__':unittest.main(verbosity=2)
