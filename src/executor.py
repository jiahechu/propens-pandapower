"""
Execute the toolbox.
"""
import pandapower as pp
from frontend import read_input


def executor(input_setup, output_setup):
    # create pandapower network from user-defined Excel file or pre-defined topology
    if input_setup['used_pre_defined_net'] is '':
        net = read_input.read_input(input_setup['loadgen_path'], input_setup['topology_path'], input_setup['output_path']+'/network.xlsx')
    else:
        pass

    # apply scenario from data


    return 0
