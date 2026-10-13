import sys,unittest
from pathlib import Path
from datetime import date
from types import SimpleNamespace
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from date_availability import available_bounds,available_comparisons,default_custom

class AvailabilityTests(unittest.TestCase):
    def test_observed_dates_do_not_expand_from_certification_or_finance(self):
        ds=SimpleNamespace(rows=[{'grain':'orders','start':date(2026,9,21),'end':date(2026,10,4)}, {'grain':'complement','start':date(2020,1,1),'end':date(2030,1,1)}],items=[])
        self.assertEqual(available_bounds(ds),(date(2026,9,21),date(2026,10,4)))
    def test_last_complete_week_uses_data_not_today(self):
        bounds=date(2026,9,21),date(2026,10,4)
        options=available_comparisons('wow',bounds)
        self.assertEqual(options[-1][1:],((date(2026,9,28),date(2026,10,4)),(date(2026,9,21),date(2026,9,27))))
        self.assertFalse(available_comparisons('mom',bounds))
    def test_short_upload_never_defaults_outside_data(self):
        bounds=date(2025,1,2),date(2025,1,4)
        self.assertTrue(all(bounds[0]<=d<=bounds[1] for d in default_custom(bounds)))
        self.assertFalse(available_comparisons('wow',bounds))
    def test_no_dates_and_items_fallback(self):
        self.assertIsNone(available_bounds(SimpleNamespace(rows=[],items=[])))
        d=date(2025,6,1)
        self.assertEqual(available_bounds(SimpleNamespace(rows=[],items=[{'start':d,'end':d}])),(d,d))
