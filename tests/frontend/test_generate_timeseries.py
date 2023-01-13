"""
Test generate_timeseries.py.
"""

import unittest
import pandapower.networks as pn
from src.frontend.generate_timeseries import generate_timeseries


class TestGenerateTS(unittest.TestCase):
    # define a network
    net = pn.case5()
    ts_path = './timeseries.xlsx'
    net, num_ts = generate_timeseries(net, ts_path)

    def test_num_ts(self):
        self.assertEqual(self.num_ts, 8761)

    def test_controllers(self):
        self.assertEqual(self.net.controller.empty, False)
        self.assertEqual(len(self.net.controller), 4)


if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()
