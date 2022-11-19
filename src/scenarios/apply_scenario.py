"""
Apply pre-defined scenario to network.
"""
import scenarios

def apply_scenario(net, scenario, **kwargs):
    """
    Apply a given pre-defined scenario to network.

    Args:
        net: pandapower network.
        scenario: name of the pre-defined scenario (string).

    Returns:
        net: pandapower network after applying the scenario.
    """
    scenario_name = 'scenario_' + scenario
    scenario_fct = getattr(scenarios, scenario_name)
    net = scenario_fct(net, **kwargs)

    return net
