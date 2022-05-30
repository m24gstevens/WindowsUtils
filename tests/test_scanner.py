import unittest
import os
from WindowsUtils.scanner import (get_file_info, scan_file,
                                  RESULT_NOT_DETECTED)

bat_path = os.path.normpath(os.path.join(__file__, "..\\bat_test.bat"))
bad_path = os.path.normpath(os.path.join(__file__, "..\\fail.txt"))

class Test(unittest.TestCase):
    def test_open(self):
        self.assertIsNotNone(get_file_info(bat_path))
        self.assertRaises(Exception, get_file_info, bad_path)

    def test_scan(self):
        scan_result = scan_file(bat_path)
        self.assertIsInstance(scan_result, dict)
        self.assertIsInstance(scan_result['Is Malware'], bool)
        self.assertIsInstance(scan_result['Risk level'], int)
        self.assertEqual(scan_result['Risk level'], RESULT_NOT_DETECTED)
        self.assertFalse(scan_result['Is Malware'])

if __name__ == "__main__":
    unittest.main()
