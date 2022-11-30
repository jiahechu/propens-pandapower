"""
Database of all pre-defined scenarios.
"""


def scenario_pv_gen(net, para):
    """
    Change all PV generation to given percent.

    Args:
        net: pandapower network.
        para: percent of changing (0-2).

    Returns:
        net: pandapower network after applying the scenario.
    """
    # change pv generation value
    if para < 0 or para > 2:
        raise ValueError('The parameter for pre-defined pv-gen scenario should between 0 and 2 (0% to 200%)')
    else:
        net.gen['p_mw'][net.gen['type'] == 'pv'] *= para

    return net


def scenario_wind_gen(net, para):
    """
    Change all wind generation to given percent.

    Args:
        net: pandapower network.
        para: percent of changing (0-2).

    Returns:
        net: pandapower network after applying the scenario.
    """
    # change wind generation value
    if para < 0 or para > 2:
        raise ValueError('The parameter for pre-defined wind-gen scenario should between 0 and 2 (0% to 200%)')
    else:
        net.gen['p_mw'][net.gen['type'] == 'wind'] *= para

    return net


def scenario_conventional_pp_gen(net, para):
    """
    Change all conventional power plant generation to given percent.

    Args:
        net: pandapower network.
        para: percent of changing (0-2).

    Returns:
        net: pandapower network after applying the scenario.
    """
    # change conv pp generation value
    if para < 0 or para > 1:
        raise ValueError('The parameter for pre-defined conv-pp scenario should between 0 and 1 (0% to 100%)')
    else:
        net.sgen['p_mw'][net.sgen['type'] == 'conv pp'] *= para

    return net


def scenario_load(net, para):
    """
    Change all load to given percent.

    Args:
        net: pandapower network.
        para: percent of changing (0-2).

    Returns:
        net: pandapower network after applying the scenario.
    """
    if para < 0:
        raise ValueError('The parameter for pre-defined load scenario should bigger than 0 (>0%)')
    else:
        net.load['p_mw'][:] *= para
    
    return net


def scenario_trafo_cap(net, para):
    # do sth
    if para < 0.5 or para > 2:
        raise ValueError('The parameter for pre-defined load scenario should between 0,5 and 2 (50% to 200%)')
    else
        net.trafo['sn_mva'][:] *= para
    return net


def scenario_lines_cap(net, para):
    # do sth
    return net


def scenario_storage(net, para):
    # do sth
    if para < 0 or para > 1:
        raise ValueError('The parameter for pre-defined load scenario should between 0 and 1 (0% to 100%)')
    else
        net.storage['max_e_mwh'][:] *= para
    return net


def scenario_switch(net, zone):
    # do sth
    return net
