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
        # static generator
        net.sgen['p_mw'][net.sgen['type'] == 'pv'] *= para
        net.sgen['q_mvar'][net.sgen['type'] == 'pv'] *= para

        # generator
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
        # static generator
        net.sgen['p_mw'][net.sgen['type'] == 'wind'] *= para
        net.sgen['q_mvar'][net.sgen['type'] == 'wind'] *= para

        # generator
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
        net.sgen['q_mvar'][net.sgen['type'] == 'conv pp'] *= para

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
        raise ValueError('The parameter for pre-defined load scenario should be bigger than 0 (>0%)')
    else:
        net.load['p_mw'][:] *= para
        net.load['q_mvar'][:] *= para
    
    return net


def scenario_trafo_cap(net, para):
    """
    Change all transformer capacities to given percent.

    Args:
        net: pandapower network.
        para: percent of changing (0-2).

    Returns:
        net: pandapower network after applying the scenario.
    """
    if para < 0.5 or para > 2:
        raise ValueError('The parameter for pre-defined trafo capacity scenario should between 0,5 and 2 (50% to 200%)')
    else:
        net.trafo['sn_mva'][:] *= para

    return net


def scenario_lines_cap(net, para):
    if para < 0:
        raise ValueError('The parameter for pre-defined lines_cap scenario should be bigger than 0 (>0%)')
    elif para > 1.95:
        print("Warning! Overload may happen! Now adding parallel lines instead of changing the thermal current ")
        net.line['parallel'][:] *= round(para, 0)
    else:
        print("Input value has impact only during optimal power flow analysis")
        net.line['max_i_ka'][:] *= para

    return net


def scenario_storage(net, para):
    # if para < 0 or para > 1:
    #     raise ValueError('The parameter for pre-defined storage scenario should between 0 and 1 (0% to 100%)')
    # else
    #     net.storage['max_e_mwh'][:] *= para
    return net


def scenario_switch(net, zone):
    # if para == 0:
    #     net.switch['closed'][:] = False
    # elif para == 1:
    #     net.switch['closed'][:] = True
    # else:
    #     raise ValueError('The parameter for pre-defined switch scenario should be 0 or 1')
    return net