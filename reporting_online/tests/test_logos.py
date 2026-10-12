import sys
import unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from online_workspace import Workspace
from PIL import Image

class LogoTests(unittest.TestCase):
    def test_local_svg_conversion_preserves_ratio(self):
        ws=Workspace()
        try:
            svg=b'<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><rect width="200" height="100" fill="#6C2BD9"/></svg>'
            path=ws.save_logo('test.svg',svg)
            with Image.open(path) as image:self.assertEqual(image.width/image.height,2)
            self.assertTrue(path.is_relative_to(ws.root))
        finally:ws.close()
    def test_external_svg_rejected(self):
        ws=Workspace()
        try:
            with self.assertRaises(ValueError):ws.save_logo('bad.svg',b'<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/image.png"/></svg>')
        finally:ws.close()

if __name__=='__main__':unittest.main(verbosity=2)
