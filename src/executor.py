"""
Execute the toolbox.
"""
from src.frontend.read_input import read_input
from src.frontend.generate_timeseries import generate_timeseries
from src.scenarios.apply_scenario import apply_scenario
from src.analysis.solver import solve
from src.analysis.excel_output import create_excel


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
    for scenario_name, scenario_path, pd_scenario, pd_para in input_setup['scenario_setup']:
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
        network_name = input_setup['topology_name']
        output_path = output_setup['output_path']
 #%%    #Solve the network depeding of it is a time_series or one iteration analysis 
        # if times_step is 1, everything is saved in net.res_, thus 'results' is empty
        results, net = solve(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps)   
        #Call the exxcel template, fill up with the results, and save the results in a new excel spreadsheet
        # if times_step is 1, everything is saved in net.res_, thus 'tables' is empty
        tables = create_excel(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps, results)  
              
    #%% 
    return 0
