"""
Test parameters.py.
"""
import unittest
import pandapower as pp
import pandapower.networks as pn
import parameters as par

def test_in(element, elements): 
    flag = element in elements
    return flag

class TestParameters(unittest.TestCase):
    net = pp.create_empty_network( name = 'test_net')
    parameters = par.select_parameters()
    sheet, cell, elements_by_type = par.sheets_parameters()
    
    def test_sheets_parameters(self):
        sheet, cell, elements_by_type = par.sheets_parameters()
        # test whether the elements exits in the pandapower network
        elements = self.net.keys()
        for element_type in elements_by_type:
            if len(elements_by_type[element_type]) > 1:
                for element in elements_by_type[element_type]:
                    self.assertEqual(test_in(element,elements), True)
            if len(elements_by_type[element_type]) == 1:                                                 
                self.assertEqual(test_in(elements_by_type[element_type][0],elements), True)
        for sheet_initial in cell: 
            self.assertEqual(test_in(cell[sheet_initial] , cell['initial']), True)                                                     
        for element in sheet:
            self.assertEqual(test_in(element,elements_by_type[sheet[element]]), True)
          
    def test_output_parameters(self):
        number, column, parameters = par.output_parameters(self.net, gen_fuel_tech = [], scenario_name = [])  
        elements = self.net.keys()
        for element in number:                                                                                    
            self.assertEqual(test_in(element,elements), True)

    def test_preallocate_table(self):
        number, column, parameters = par.output_parameters(self.net, gen_fuel_tech = [], scenario_name = [])  
        table_df = par.preallocate_table('bus', column, number)
        self.assertEqual(list(table_df.keys()), column['parameter']['bus'])
    
    def test_add_fuel(self):
        net = pn.case5()
        gen_fuel_tech = []
        sheet, cell, elements_by_type = par.sheets_parameters()
        net = par.add_fuel(net,gen_fuel_tech, elements_by_type)
        for element in elements_by_type['Generation']:
            self.assertIn('fuel',list(net[element].keys()))
    
if __name__ == '__main__':
    # begin the unittest.main()
    unittest.main()
