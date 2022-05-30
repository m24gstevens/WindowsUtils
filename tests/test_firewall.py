import unittest
import os
from WindowsUtils import firewall
import comtypes

from ctypes import POINTER, oledll

notepad_path = os.environ['WINDIR'] + "\\system32\\notepad.exe"

class Test(unittest.TestCase):         
    def test_firewall_settings(self):
        getters_setters = [
            (firewall.get_firewall_enabled, firewall.put_firewall_enabled),
            (firewall.get_block_inbound, firewall.put_block_inbound),
            (firewall.get_notifications_disabled, firewall.put_notifications_disabled),
            (firewall.get_unicast_response_to_multicast, firewall.put_unicast_response_to_multicast),
            (firewall.get_default_inbound_block, firewall.put_default_inbound_block),
            (firewall.get_default_outbound_block, firewall.put_default_outbound_block)
            ]
        for profile in (0x1, 0x2, 0x4):     #3 possible profiles
            for getter, setter in getters_setters:
                init_value = getter(profile)
                self.assertIsInstance(init_value, bool)
                self.assertIsNone(setter(True, profile))
                self.assertIsInstance(getter(profile), bool)
                self.assertTrue(getter(profile))
                self.assertIsNone(setter(False, profile))
                self.assertIsInstance(getter(profile), bool)
                self.assertFalse(getter(profile))
                setter(init_value, profile)

    def test_add_rule(self):
        firewall.add_rule("WindowsUtils test", "Filler description",
                          firewall.IP_PROTOCOL_TCP,
                          True, True, direction = firewall.OUTBOUND,
                          application = notepad_path,
                          interface_types = "Wireless", local_port = 60000,
                          remote_port = 60000)
        policy_itf = firewall._init_settings()
        self.assertIsInstance(policy_itf, comtypes.POINTER(firewall.INetFwPolicy2))
        rules = policy_itf.get_Rules()
        self.assertIsInstance(rules, comtypes.POINTER(firewall.INetFwRules))
        self.assertIsInstance(rules.get_Count(), int)
        self.assertGreater(rules.get_Count(), 0)
        rule = rules.Item("WindowsUtils test")
        self.assertIsInstance(rule, comtypes.POINTER(firewall.INetFwRule))
        self.assertTrue(bool(rule))
        self.assertIsInstance(rule.get_Description(), str)
        self.assertEqual(rule.get_ApplicationName(), notepad_path)
        self.assertEqual(rule.get_Protocol(), firewall.IP_PROTOCOL_TCP)
        self.assertEqual(rule.get_LocalPorts(), '60000')
        self.assertIsInstance(rule.get_InterfaceTypes(), str)

if __name__ == "__main__":
    unittest.main()   
        
    
