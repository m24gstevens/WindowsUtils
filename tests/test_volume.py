import unittest
from ctypes import POINTER

from WindowsUtils.volume import AudioEndpoint, IAudioEndpointVolume

class Test(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.Volume = AudioEndpoint()
        cls.Volume.__enter__()
        
    @classmethod
    def tearDownClass(cls):
        cls.Volume.__exit__()
        del cls.Volume

    def test_init(self):
        self.assertIsInstance(self.Volume._endpt, POINTER(IAudioEndpointVolume))

    def test_mute(self):
        self.assertIsNone(self.Volume.set_mute(True))
        self.Volume.set_mute(True)
        self.assertTrue(self.Volume.get_mute())
        self.Volume.set_mute(False)
        self.assertFalse(self.Volume.get_mute())

    def test_get_set(self):
        self.assertIsInstance(self.Volume.get_master_volume(), float)
        self.assertIsNone(self.Volume.set_master_volume(0.70))
        for level in (0.02, 0.10, 0.50, 0.75, 0.98):
            self.Volume.set_master_volume(level)
            self.assertEqual(round(100*(self.Volume.get_master_volume() - level)), 0)

    def test_increments(self):
        init_volume = self.Volume.get_master_volume()
        self.assertIsNone(self.Volume.increment_master_volume())
        self.assertIsNone(self.Volume.decrement_master_volume())
        self.Volume.set_master_volume(0.7)
        for i in range(5):
            self.Volume.increment_master_volume()
        for i in range(5):
            self.Volume.decrement_master_volume()
        self.assertEqual(round(100*(self.Volume.get_master_volume() - 0.7)),0)
        self.Volume.set_master_volume(init_volume)

    def test_ranges(self):
        self.assertIsInstance(self.Volume.get_range(), tuple)
        self.assertEqual(len(self.Volume.get_range()), 3)
        for i in range(3):
            self.assertIsInstance(self.Volume.get_range()[i], float)
            

if __name__ == "__main__":
    unittest.main()
        
