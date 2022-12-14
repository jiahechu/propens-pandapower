"""
Test read_input.py.
Only test bus, line, generation and load.
"""

import unittest
from read_input import read_input


class TestReadInput(unittest.TestCase):
    net, ts_setup = read_input('./pv_storage.xlsx', './topology.xlsx')

    def test_buses(self):
        buses = self.net.bus
        self.assertListEqual(buses['name'], ['Trafostation_OS', 'main_busbar', 'bus_1_1', 'bus_1_2', 'bus_1_3', 'bus_1_4',
                                         'bus_1_5', 'bus_1_6', 'bus_2_1', 'bus_2_2'])
        self.assertListEqual(buses['vn_kv'], [10, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4])

    def test_lines(self):
        lines = self.net.line
        self.assertEqual(lines['std_type'], ['NFA2X 4x70']*8)
        self.assertEqual(lines['from_bus'], range(1, 8))
        self.assertEqual(lines['to_bus'], range(2, 9))


if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()
