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
