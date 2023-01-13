"""
Execute the toolbox.
"""
from src.frontend.read_input import read_input
from src.frontend.generate_timeseries import generate_timeseries
from src.scenarios.apply_scenario import apply_scenario
from src.analysis.solver import solve
from src.analysis.excel_output import create_excel


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
    for scenario_name, scenario_path, pd_scenario, pd_para in input_setup['scenario_setup']:
        print('\nStart simulation with scenario ' + scenario_name)

        """Frontend"""
        # create pandapower network from Excel
        try:
            print('\nReading input excel files')
            net, general, gen_fuel_tech = read_input(scenario_path, input_setup['topology_path'])
            time_steps = 1
            print('> Done')
        except:
            print('\nError while reading excel files, please check input file again!')
            print('Program stops.')
            print('Detail error arguments: ')
            raise

        # apply scenario from data
        if pd_scenario != '':
            try:
                print('\nApplying pre-defined scenarios')
                net = apply_scenario(net, pd_scenario, pd_para)
                print('> Done')
            except:
                print('\nError while applying pre-defined scenarios, please check input parameters again!')
                print('Program stops.')
                print('Detail error arguments: ')
                raise

        # apply time series
        if general['use_ts'][0]:
            try:
                print('\nGenerating controllers for time series analysis')
                net, time_steps = generate_timeseries(net, general['ts_path'][0])
                print('> Done')
            except:
                print('\nError while generating controllers for time series analysis, please check time series data again!')
                print('Program stops.')
                print('Detail error arguments: ')
                raise

        print('\nSuccessfully read pandapower network from Excel:')
        print(net)

        """Analysis"""
        # parameters to define the output file name, and its path
        gen_fuel_tech = []  # to be read
        # if times_step is 1, everything is saved in net.res_, thus 'results' is empty
        try:
            results, net = solve(input_setup['topology_name'], scenario_name, gen_fuel_tech, output_setup['output_path'],
                                 net, time_steps)
        except:
            print('\nError while solving network, e.g. not converging')
            print('Program stops.')
            print('Detail error arguments: ')
            raise
        # Call the excel template, fill up with the results, and save the results in a new excel spreadsheet
        # if times_step is 1, everything is saved in net.res_, thus 'tables' is empty
        try:
            tables = create_excel(input_setup['topology_name'], scenario_name, gen_fuel_tech, output_setup['output_path'],
                                  net, time_steps, results)
        except:
            print('\nError while creating output excel file')
            print('Program stops.')
            print('Detail error arguments: ')
            raise

        # optimal power flow
        if general['use_opf'][0]:
            try:
                print('\nCalculating optimal power flow')
                pass  # do opf analysis
            except:
                print('\nError while doing optimal power flow')
                print('Program stops.')
                print('Detail error arguments: ')
                raise
