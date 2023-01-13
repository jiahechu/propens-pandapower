"""
Test read_input.py.
Only test bus, line, switch, generation, trafo and load.
"""

import unittest
from src.frontend.read_input import read_input


class TestReadInput(unittest.TestCase):
    net, general, fuel = read_input('./pv_storage.xlsx', './topology.xlsx')

    def test_buses(self):
        buses = self.net.bus
        self.assertEqual(buses['name'].values.tolist(), ['Trafostation_OS', 'main_busbar', 'bus_1_1', 'bus_1_2',
                                                         'bus_1_3', 'bus_1_4', 'bus_1_5', 'bus_1_6', 'bus_2_1',
                                                         'bus_2_2'])
        self.assertEqual(buses['vn_kv'].values.tolist(), [10, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4])

    def test_lines(self):
        lines = self.net.line
        self.assertEqual(lines['std_type'].values.tolist(), ['NFA2X 4x70']*8)
        self.assertEqual(lines['from_bus'].values.tolist(), [1, 2, 3, 4, 5, 6, 1, 8])
        self.assertEqual(lines['to_bus'].values.tolist(), [2, 3, 4, 5, 6, 7, 8, 9])
        self.assertEqual(lines['parallel'].values.tolist(), [1]*8)

    def test_switches(self):
        switches = self.net.switch
        self.assertEqual(switches['bus'].values.tolist(), [2, 3, 4, 5, 6, 7, 8, 9])
        self.assertEqual(switches['element'].values.tolist(), [2, 3, 4, 5, 6, 7, 8, 9])
        self.assertEqual(switches['et'].values.tolist(), ['b']*8)

    def test_loads(self):
        loads = self.net.load
        self.assertEqual(loads['bus'].values.tolist(), [2, 3, 4, 5, 6, 7, 8, 9])
        self.assertEqual(loads['p_mw'].values.tolist(), [0.008]*8)
        self.assertEqual(loads['q_mvar'].values.tolist(), [0] * 8)
        self.assertEqual(loads['type'].values.tolist(), ['wye']*8)

    def test_sgen(self):
        sgens = self.net.sgen
        self.assertEqual(sgens['bus'].values.tolist(), [2, 3, 4, 5, 6, 7, 8, 9])
        self.assertEqual(sgens['p_mw'].values.tolist(), [0.006] * 8)
        self.assertEqual(sgens['q_mvar'].values.tolist(), [0] * 8)
        self.assertEqual(sgens['type'].values.tolist(), ['pv'] * 8)

    def test_trafo(self):
        trafo = self.net.trafo
        self.assertEqual(trafo['hv_bus'].values.tolist(), [0])
        self.assertEqual(trafo['lv_bus'].values.tolist(), [1])
        self.assertEqual(trafo['std_type'].values.tolist(), ['0.1 MVA 10/0.4 kV'])

    def test_general(self):
        use_ts = self.general['use_ts']
        ts_path = self.general['ts_path']
        use_opf = self.general['use_opf']
        self.assertEqual(use_ts.values.tolist(), [True])
        self.assertEqual(ts_path.values.tolist(), ['./tests/frontend/timeseries.xlsx'])
        self.assertEqual(use_opf.values.tolist(), [False])

    def test_fuel(self):
        gen_type = self.fuel['gen_type']
        index = self.fuel['index']
        fuel = self.fuel['fuel']
        self.assertEqual(gen_type.values.tolist(), ['sgen'] * 8)
        self.assertEqual(index.values.tolist(), [0, 1, 2, 3, 4, 5, 6, 7])
        self.assertEqual(fuel.values.tolist(), ['solar'] * 8)


if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()
