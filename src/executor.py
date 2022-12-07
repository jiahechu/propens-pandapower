"""
Execute the toolbox.
"""
from read_input import read_input
from apply_scenario import apply_scenario
from generate_timeseries import generate_timeseries


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
        # create pandapower network from Excel
        net, ts_setup = read_input(scenario_path, input_setup['topology_path'])

        # TODO: apply pre-defined scenarios first or time series first?
        # apply scenario from data
        if pd_scenario != '':
            net = apply_scenario(net, pd_scenario, pd_para)

        # apply time series
        if ts_setup['use_ts'][0]:
            generate_timeseries(net, ts_setup['ts_path'][0])

        print(net)

    return 0
