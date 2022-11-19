"""
Execute the toolbox.
"""
from frontend import read_input
from scenarios import apply_scenario


def executor(input_setup, output_setup):
    # create pandapower network from user-defined Excel file or pre-defined topology
    if input_setup['used_pre_defined_net'] is '':
        net = read_input.read_input(input_setup['loadgen_path'], input_setup['topology_path'], output_setup['output_path'])
    else:
        pass

    # apply scenario from data
    if input_setup['used_scenario'] != '':
        net = apply_scenario.apply_scenario(net, input_setup['used_scenario'], **input_setup['scenario_para'])

    return 0
