"""
Test generate_timeseries.py.
Only test bus, line, switch, generation, trafo and load.
"""

import unittest
import pandapower.networks as pn
from generate_timeseries import generate_timeseries


class TestGenerateTS(unittest.TestCase):
    # define a network
    net = pn.case5()
    ts_path = './tests/frontend/timeseries.xlsx'
    net, num_ts = generate_timeseries(net, ts_path)


if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()