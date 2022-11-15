"""
Read network data from Excel file.
"""

import pandapower as pp
import pandas as pd


def read_input(file_path):
    """
    Read network data from Excel file.

    Args:
        file_path: path of the Excel file.

    Returns:
        net: pandapower network.
    """
    # init
    net = pp.create_empty_network()
    net_components = ['bus', 'line', 'switch', 'load', 'motor', 'asymmetric_load', 'sgen', 'ext_grid', 'transformer',
                      'gen', 'impedance', 'dcline', 'storage']

    for comp in net_components:
        # read parameters from Excel
        data = pd.read_excel(file_path, sheet_name=comp, index_col=0).T.to_dict()

        # add components to network
        if comp == 'line':  # special case for line (std or user-defined)
            for key in data.keys():
                if data[key]['std_type'] == '':     # create pre-defined line from standard library
                    data[key].pop('std_type')
                    pp.create_line_from_parameters(net, **data[key])
                else:   # create user-defined line from parameters
                    for k in ['r_ohm_per_km', 'x_ohm_per_km', 'c_nf_per_km', 'g_us_per_km', 'r0_ohm_per_km',
                              'x0_ohm_per_km', 'c0_nf_per_km', 'max_i_ka', 'type']:
                        data[key].pop(k)
                    pp.create_line(net, **data[key])
        elif comp == 'transformer':     # special case for trafo (std or user-defined)
            for key in data.keys():
                if data[key]['std_type'] == '':     # create pre-defined trafo from standard library
                    data[key].pop('std_type')
                    pp.create_transformer_from_parameters(net, **data[key])
                else:   # create user-defined trafo from parameters
                    for k in ['sn_mva', 'vn_hv_kv', 'vn_lv_kv', 'vk_percent', 'vkr_percent', 'pfe_kw', 'i0_percent',
                              'shift_degree', 'tap_side', 'tap_neutral', 'tap_min', 'tap_max', 'tap_step_percent',
                              'tap_step_degree', 'tap_phase_shifter']:
                        data[key].pop(k)
                    pp.create_transformer(net, **data[key])
        else:   # all other components
            fct_name = 'create_' + comp
            fct = getattr(pp, fct_name)     # get create function from pandapower
            for key in data.keys():
                fct(net, **data[key])

    return net


net = read_input("./template_input.xlsx")
