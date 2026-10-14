import json
import sys
import unittest
from datetime import date
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from online_workspace import Workspace

class WorkspaceTests(unittest.TestCase):
    def setUp(self):self.a=Workspace();self.b=Workspace()
    def tearDown(self):self.a.close();self.b.close()
    def test_isolation(self):
        self.a.load_demo()
        self.assertNotEqual(self.a.root,self.b.root);self.assertEqual(len(self.a.dataset.rows),417);self.assertEqual(self.b.dataset.rows,[])
        self.assertGreater(len(self.a.dataset.items),0)
    def test_config_roundtrip_and_recognition(self):
        self.a.load_demo();raw=self.a.config_bytes(['DEMO_BELHAJ'],{'mode':'Deux périodes libres'})
        exported=json.loads(raw)
        self.assertNotIn(str(self.a.root),raw.decode());self.assertNotIn('APP_PASSWORD',raw.decode())
        self.b.apply_config(raw)
        uploads=[(s['name'],Path(s['path']).read_bytes()) for s in self.a.sources]
        errors,n=self.b.add_uploads(uploads)
        self.assertEqual(errors,[]);self.assertEqual(n,3)
        self.assertTrue(any(c['grain']=='items' for c in self.b.dataset.coverage));self.assertEqual(self.b.settings,self.a.settings)
    def test_changed_file_requires_recertification(self):
        self.a.load_demo();self.b.apply_config(self.a.config_bytes())
        source=next(s for s in self.a.sources if s['name']=='orders_DEMO.csv')
        raw=Path(source['path']).read_bytes()+b'\n'
        self.b.add_uploads([(source['name'],raw)])
        self.assertFalse(self.b.sources[0]['sheets'][0]['complete'])
    def test_config_server_paths_not_imported(self):
        raw=json.dumps({'schema_version':1,'settings':{},'theme':{'yassir_logo':'/etc/passwd','restaurant_logos':{'x':'/etc/passwd'}},'source_configs':[]}).encode()
        self.a.apply_config(raw);self.assertEqual(self.a.theme['yassir_logo'],'');self.assertEqual(self.a.theme['restaurant_logos'],{})
    def test_export_is_scoped(self):
        self.a.load_demo();interval=(date(2026,9,28),date(2026,10,4));previous=(date(2026,9,21),date(2026,9,27))
        result=self.a.generate(['DEMO_BELHAJ','DEMO_RABAT'],interval,previous)
        self.assertTrue(Path(result['zip']).is_relative_to(self.a.root));self.assertEqual(len(result['results']),2);self.assertEqual(self.b.results,[])
        from pptx import Presentation
        self.assertEqual(len(Presentation(result['results'][0]['pptx']).slides),20)
    def test_reset_only_own_session(self):
        self.a.load_demo();self.b.load_demo();self.a.reset();self.assertEqual(self.a.dataset.rows,[]);self.assertEqual(len(self.b.dataset.rows),417)
    def test_failed_config_keeps_settings(self):
        self.a.settings={'tax_basis':'TTC'}
        with self.assertRaises(ValueError):self.a.apply_config(json.dumps({'schema_version':1,'settings':{},'theme':{},'source_configs':[{'sheets':[{'grain':'UNKNOWN'}]}]}).encode())
        self.assertEqual(self.a.settings,{'tax_basis':'TTC'})

if __name__=='__main__':unittest.main(verbosity=2)
