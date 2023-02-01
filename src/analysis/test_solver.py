# -*- coding: utf-8 -*-
"""
Test solver.py.

"""
import unittest
import pandapower as pp
import pandapower.networks as pn
import solver
# import parameters as par

def create_opf_net():
    net = pp.create_empty_network()
    
    min_vm_pu = .95
    max_vm_pu = 1.05
    
    # create buses
    bus1 = pp.create_bus(net, vn_kv=110., min_vm_pu=min_vm_pu, max_vm_pu=max_vm_pu)
    bus2 = pp.create_bus(net, vn_kv=110., min_vm_pu=min_vm_pu, max_vm_pu=max_vm_pu)
    bus3 = pp.create_bus(net, vn_kv=110., min_vm_pu=min_vm_pu, max_vm_pu=max_vm_pu)
    
    # create 110 kV lines
    pp.create_line(net, bus1, bus2, length_km=1., std_type='149-AL1/24-ST1A 110.0')
    pp.create_line(net, bus2, bus3, length_km=1., std_type='149-AL1/24-ST1A 110.0')
    pp.create_line(net, bus3, bus1, length_km=1., std_type='149-AL1/24-ST1A 110.0')
    
    # create loads
    pp.create_load(net, bus3, p_mw=300)
    
    # create generators
    g1 = pp.create_gen(net, bus1, p_mw=200, min_p_mw=0, max_p_mw=300, controllable=True, slack=True)
    g2 = pp.create_gen(net, bus2, p_mw=0, min_p_mw=0, max_p_mw=300, controllable=True)
    g3 = pp.create_gen(net, bus3, p_mw=0, min_p_mw=0, max_p_mw=300, controllable=True)
    
    pp.create_poly_cost(net, element=g1, et="gen", cp1_eur_per_mw=30)
    pp.create_poly_cost(net, element=g2, et="gen", cp1_eur_per_mw=30)
    pp.create_poly_cost(net, element=g3, et="gen", cp1_eur_per_mw=29.999)
    
    return net

class TestSolver(unittest.TestCase):   
    # parameters = par.select_parameters()
    # sheet, cell, elements_by_type = par.sheets_parameters()
    
    def test_solve(self):
        network_name = 'test'
        scenario_name = '_solver'
        gen_fuel_tech = []
        
        folder_name = 'test_data'
        output_path = './'+ folder_name
        
        net = pn.case5()
        time_steps = '1'
        general = {}
        general['use_ts'] = [False, False]
        general['use_opf'] = [False, False]
        general['use_dc'] = [False, False]        
        results, net = solver.solve(network_name, scenario_name, gen_fuel_tech, output_path, net, time_steps, general)
        for element in results:
            self.assertEqual(results[element], [])
        
    def test_run_one_iteration_PF_AC(self):
        network_name = 'test_PF'
        scenario_name = 'AC_run_one_iteration'
        gen_fuel_tech = []
        
        folder_name = 'test_data'
        output_path = './'+ folder_name

        net = pn.case5()
        # time_steps = '1'
        general = {}
        general['use_ts'] = [False, False]
        general['use_opf'] = [False, False]
        general['use_dc'] = [False, False]
        results, net = solver.run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, general)
        if len(net['res_bus']) > 0: flag = True
        self.assertEqual(flag, True)
        
    def test_run_one_iteration_PF_DC(self):
        network_name = 'test_PF'
        scenario_name = 'DC_run_one_iteration'
        gen_fuel_tech = []
        
        folder_name = 'test_data'
        output_path = './'+ folder_name

        net = pn.case5()
        # time_steps = '1'
        general = {}
        general['use_ts'] = [False, False]
        general['use_opf'] = [False, False]
        general['use_dc'] = [True, True]
        results, net = solver.run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, general)
        if len(net['res_bus']) > 0: flag = True
        self.assertEqual(flag, True) 
    
    def test_run_one_iteration_OPF_AC(self):
        network_name = 'test_PF'
        scenario_name = 'AC_run_one_iteration'
        gen_fuel_tech = []
        
        folder_name = 'test_data'
        output_path = './'+ folder_name

        net = create_opf_net()
        # time_steps = '1'
        general = {}
        general['use_ts'] = [False, False]
        general['use_opf'] = [True, True]
        general['use_dc'] = [False, False]
        results, net = solver.run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, general)
        if len(net['res_bus']) > 0: flag = True
        self.assertEqual(flag, True)
        
    def test_run_one_iteration_OPF_DC(self):
        network_name = 'test_PF'
        scenario_name = 'DC_run_one_iteration'
        gen_fuel_tech = []
        
        folder_name = 'test_data'
        output_path = './'+ folder_name

        net = create_opf_net()
        # time_steps = '1'
        general = {}
        general['use_ts'] = [False, False]
        general['use_opf'] = [True, True]
        general['use_dc'] = [True, True]
        results, net = solver.run_one_iteration(network_name, scenario_name, gen_fuel_tech, output_path, net, general)
        if len(net['res_bus']) > 0: flag = True
        self.assertEqual(flag, True) 
        

if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()

