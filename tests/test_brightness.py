import unittest
import os
from WindowsUtils import brightness
import time

from ctypes.wintypes import HANDLE

class Test(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.monitors = brightness.get_monitors()
        cls.wmi_monitors = list(filter(lambda mon: type(mon) == brightness.WmiMonitor,
                                        cls.monitors))
        cls.vcp_monitors = list(filter(lambda mon: type(mon) == brightness.VCPMonitor,
                                        cls.monitors))

    def test_wmi_setup(self):
        for mon in self.wmi_monitors:
            self.assertIsInstance(mon.system_name, str)
            self.assertIsInstance(mon.displayed_name, str)
            self.assertIsInstance(mon.manufacturer_code, str)
            self.assertEqual(len(mon.manufacturer_code), 3)

    def test_vcp_setup(self):
        for mon in self.vcp_monitors:
            self.assertIsInstance(mon.system_name, str)
            self.assertNotIn('_', mon.system_name)
            self.assertIsInstance(mon.displayed_name, str)
            self.assertIsInstance(mon.manufacturer_code, str)
            self.assertEqual(len(mon.manufacturer_code), 3)
            self.assertIsInstance(mon._display_handle, HANDLE)

    def _test_brightness(self, lst):
        for mon in lst:
            init_brightness = mon.get_brightness()
            self.assertIsInstance(init_brightness, int)
            for b in (-80, 0, 20, 40, 80, 100, 120):
                self.assertIsNone(mon.set_brightness(b))
                time.sleep(1)
                self.assertIsInstance(mon.get_brightness(), int)
                self.assertEqual(mon.get_brightness(), max(min(b,100),0))
            mon.set_brightness(init_brightness)

    def test_wmi_brightness(self):
        self._test_brightness(self.wmi_monitors)

    def test_vcp_brightness(self):
        self._test_brightness(self.vcp_monitors)


if __name__ == "__main__":
    unittest.main()
            
            
        
