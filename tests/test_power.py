import unittest
import os
from WindowsUtils import power
from WindowsUtils.helpers.powertypes import *
import time

class Test(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.initial_scheme = power.PowerScheme.active_scheme()

    @classmethod
    def tearDownClass(cls):
        cls.initial_scheme.set_this_scheme()
        
    def test_profiles(self):
        schemes = power.get_all_schemes()
        for scm in schemes:
            self.assertIsInstance(scm, power.PowerScheme)
            self.assertIsInstance(scm.name, str)
            self._scheme_test(scm)

    def _scheme_test(self, scm):
        for bl in (True, False):
            init_sleep = scm.get_sleep(bl)
            self.assertIsInstance(init_sleep, int)
            init_sleep = scm.get_sleep(bl)
            for tst in (-10, 0, 100, 400, 8000):
                self.assertIsNone(scm.sleep_after(bl,tst))
                scm.sleep_after(bl,tst)
                self.assertEqual(scm.get_sleep(bl), max(0,tst))
            scm.sleep_after(bl,init_sleep)

            init_act_cool = scm.get_active_cooling(bl)
            self.assertIsInstance(init_act_cool, int)
            self.assertIsNone(scm.set_active_cooling(bl,True))
            scm.set_active_cooling(bl,True)
            self.assertTrue(scm.get_active_cooling(bl))
            self.assertIsNone(scm.set_active_cooling(bl,False))
            scm.set_active_cooling(bl,False)
            self.assertFalse(scm.get_active_cooling(bl))
            scm.set_active_cooling(bl,init_act_cool)

            self.assertRaises(Exception, scm.read_value, bl, "TEST", "TEST")
            for tst in (True, 0, None, 1.8):
                self.assertRaises(Exception, scm.set_value, bl, "TEST", "TEST", tst)

    def test_switch_schemes(self):
        self.assertIsNone(power.PowerScheme.MaxPowerSavings().set_this_scheme())
        self.assertIsNone(power.PowerScheme.MinPowerSavings().set_this_scheme())
        self.assertIsNone(power.PowerScheme.TypicalPowerSavings().set_this_scheme())

if __name__ == "__main__":
    unittest.main()
