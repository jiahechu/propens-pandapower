"""
Apply pre-defined scenario to network.
"""
from src.scenarios import scenarios


def apply_scenario(net, scenario, para):
    """
    Apply a given pre-defined scenario to network.

    Args:
        net: pandapower network.
        scenario: name of the pre-defined scenario (string).
        para: scenario parameter.

    Returns:
        net: pandapower network after applying the scenario.
    """
    scenario_name = 'scenario_' + scenario
    scenario_fct = getattr(scenarios, scenario_name)
    net = scenario_fct(net, para)

    return net
