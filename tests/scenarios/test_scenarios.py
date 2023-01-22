"""
Test scenarios.py and apply_scenarios.py
"""

import unittest
import pandapower.networks as pn
import pandapower as pp
from src.scenarios.apply_scenario import apply_scenario


class TestScenarios(unittest.TestCase):
    net = pn.case5()    # choose an example pp network

    # write test functions for all scenarios
    def test_pv_gen(self):
        pp.create_sgen(self.net, bus=0, p_mw=1, q_mvar=1, type='pv')   # assign sgen to pv to test pv scenario
        net = apply_scenario(self.net, 'pv_gen', 0.5)   # apply scenario
        self.assertEqual(net.sgen['p_mw'][net.sgen['type'] == 'pv'].values.tolist(), [0.5])   # compare p
        self.assertEqual(net.sgen['q_mvar'][net.sgen['type'] == 'pv'].values.tolist(), [0.5])  # compare q

    def test_wind_gen(self):
        pp.create_sgen(self.net, bus=0, p_mw=1, q_mvar=1, type='wind')  # assign sgen to pv to test pv scenario
        net = apply_scenario(self.net, 'wind_gen', 0.5)   # apply scenario
        self.assertEqual(net.sgen['p_mw'][net.sgen['type'] == 'wind'].values.tolist(), [0.5])   # compare p
        self.assertEqual(net.sgen['q_mvar'][net.sgen['type'] == 'wind'].values.tolist(), [0.5])  # compare q
   def test_pp_gen(self):
        pp.create_sgen(self.net, bus=0, p_mw=1, q_mvar=1, type='conv pp')   # assign sgen to pp to test pp scenario
        net = apply_scenario(self.net, 'conventional_pp_gen', 0.5)   # apply scenario
        self.assertEqual(net.sgen['p_mw'][net.sgen['type'] == 'conv pp'].values.tolist(), [0.5])   # compare p
        self.assertEqual(net.sgen['q_mvar'][net.sgen['type'] == 'conv pp'].values.tolist(), [0.5])  # compare q

    def test_load(self):
        pp.create_load(self.net, bus=0, p_mw=1, q_mvar=1)  # assign load to test load scenario
        net = apply_scenario(self.net, 'load', 0.5)   # apply scenario
        self.assertEqual(net.load['p_mw'].values.tolist(), [150.0, 150.0, 200.0, 0.5])   # compare p
        self.assertEqual(net.load['q_mvar'].values.tolist(), [49.305, 49.305, 65.735, 0.5])  # compare q


    def test_trafo(self):
        pp.create_transformer(self.net, hv_bus=0, lv_bus=1, name="trafo 1", std_type="160 MVA 380/110 kV")  # assign trafo_cap to trafo_cap scenario
        net = apply_scenario(self.net, 'trafo_cap', 0.5)   # apply scenario
        self.assertEqual(net.trafo['sn_mva'].values.tolist(), [80.0])   # compare mva

    def test_line(self):
        pp.create_line(self.net, from_bus=0, to_bus=1, length_km=0.1, std_type="NAYY 4x50 SE")
        net = apply_scenario(self.net, 'lines_cap', 2.9)   # apply scenario
        self.assertEqual(net.line['parallel'].values.tolist(), [3, 3, 3, 3, 3, 3, 3])   # compare mva
        self.assertEqual(net.line['max_i_ka'].values.tolist(), [1.00408742467761, 99999.0, 99999.0, 99999.0, 99999.0, 0.60245245480657, 0.142])

    def test_storage(self):
        pp.create_storage(self.net, 1, p_mw=-30, max_e_mwh=60, soc_percent=1.0, min_e_mwh=5)
        net = apply_scenario(self.net, 'storage', 0.5)   # apply scenario
        self.assertEqual(net.storage['max_e_mwh'].values.tolist(), [30.0])   # compare mva
