"""
Execute the toolbox.
"""
from read_input import read_input
from apply_scenario import apply_scenario


def executor(input_setup, output_setup):
    """
    Execute the toolbox.

    Args:
        input_setup: dictionary contains input setups.
        output_setup: dictionary contains output setups.

    Returns:
        None
    """
    # simulate for each scenario
    for scenario_path, pd_scenario, pd_para in input_setup['scenario_setup']:
        # create pandapower network Excel
        net, ts_setup = read_input(scenario_path, input_setup['topology_path'])

        # apply scenario from data
        if pd_scenario != '':
            net = apply_scenario(net, pd_scenario, pd_para)


        print(net)

    return 0
