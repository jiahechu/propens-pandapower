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
