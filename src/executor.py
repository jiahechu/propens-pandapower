"""
Execute the toolbox.
"""
from src.frontend.read_input import read_input
from src.frontend.generate_timeseries import generate_timeseries
from src.scenarios.apply_scenario import apply_scenario
from src.analysis.time_series_func import run_time_series
from src.analysis.excel_output import run_one_iteration

# %%
def executor(input_setup, output_setup):
    # %%
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
        net, ts_setup, gen_fuel_tech = read_input(scenario_path, input_setup['topology_path'])
        time_steps = 1

        # apply scenario from data
        if pd_scenario != '':
            net = apply_scenario(net, pd_scenario, pd_para)

        # apply time series
        if ts_setup['use_ts'][0]:
            net, time_steps = generate_timeseries(net, ts_setup['ts_path'][0])

        print(net)
        # TODO: analysis once for all scenarios or once for one?
        # parameters to define the output file name, and its path
        network_name = output_setup['topology_name']
        scenario_name = output_setup['scenario_name']
        output_path = output_setup['output_path']
        
        gen_fuel_tech = [] # ------------------------------- to be readed --------------------------------
        if time_steps > 1:    
            run_time_series(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)
        else:
            run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)
        
            
    #%% 
    return 0
