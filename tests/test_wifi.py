import unittest
from ctypes import POINTER
from ctypes.wintypes import HANDLE

from WindowsUtils import wifi
from WindowsUtils.helpers.wifitypes import _DOT11_SSID, _WLAN_INTERFACE_INFO

class TestSession(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.Session = wifi.WifiSession()
        cls.Session.__enter__()

    @classmethod
    def tearDownClass(cls):
        cls.Session.__exit__()
        del cls.Session

    def test_handle(self):
        self.assertIsInstance(self.Session._use_handle, HANDLE)

    def test_security_desc(self):
        for i in range(17):
            ret = self.Session.get_security_settings(i)
            self.assertIsInstance(ret, tuple)
            self.assertEqual(len(ret), 3)
            for info in ret:
                self.assertIsInstance(info, str)

    def test_notifications(self):
        self.assertIsNone(self.Session.register_notifications())
    

class TestInterfaces(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.Session = wifi.WifiSession()
        cls.Session.__enter__()
        cls.interfaces = cls.Session.enum_interfaces()

    @classmethod
    def tearDownClass(cls):
        cls.Session.__exit__()
        del cls.Session

    def test_interfaces(self):
        for i in self.interfaces:
            self.assertIsInstance(i, wifi.WlanInterface)
            self.assertIsInstance(i._guid, wifi.GUID)
            self.assertIsInstance(i.state, int)
            self.assertIsInstance(i.description, str)

    def test_conn_attributes(self):
        for i in self.interfaces:
            self.assertIsNotNone(i.query_connection_attributes())
            conn_info = i.query_connection_attributes()
            self.assertIsInstance(conn_info, dict)
            self.assertIn('Interface state', conn_info)
            self.assertIsInstance(conn_info['Interface state'], str)
            if conn_info['Interface state'] == 'connected':   #interface on
                self.assertIn('Connection mode', conn_info)
                self.assertIsInstance(conn_info['Connection mode'], str)
                if conn_info['Connection mode'] == 'profile':
                    self.assertIn('Profile name', conn_info)
                    self.assertIsInstance(conn_info['Profile name'], str)

    def test_networks(self):
         for i in self.interfaces:
            self.assertIsNotNone(i.get_available_network_list())
            ret = i.get_available_network_list()
            for n in ret:
                self.assertIsInstance(n, wifi.AvailableNetwork)
                self.assertIsInstance(n.name, str)
                self.assertIsInstance(n.is_connectable, bool)
                self.assertIsInstance(n.bss_type, int)
                self.assertIsInstance(n.dot11_ssid, POINTER(_DOT11_SSID))
                self._connect_test(i, n)

    def _connect_test(self, interface, network):
        if network.is_connectable:
            self.assertIsNone(interface.connect(network))
            self.assertRaises(Exception, interface.connect, None)

    def test_profiles(self):
        for i in self.interfaces:
            self.assertIsNotNone(i.get_profile_list())
            ret = i.get_profile_list()
            for p in ret:
                self.assertIsInstance(p, str)
                self.assertIsInstance(i.get_profile(p)[0], str)
                self.assertRaises(Exception, i.get_profile, 'h'*129)           

if __name__ == "__main__":
    unittest.main()
        
