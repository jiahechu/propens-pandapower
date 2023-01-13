"""
Test scenarios.py and apply_scenarios.py
"""

import unittest
import pandapower.networks as pn
from src.scenarios.apply_scenario import apply_scenario


class TestScenarios(unittest.TestCase):
    net = pn.case5()

    # write test functions for all scenarios
