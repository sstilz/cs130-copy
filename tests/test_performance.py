import context 
import cProfile
import pstats
import io
import unittest
import timeit
from sheets.cell_error_type import CellErrorType, CellError
from sheets import Workbook  # Replace 'your_module' with the actual module containing the Workbook class

class PerformanceNode():
    """Class for abstractly creating graphs of cell dependencies
    
       Works as an abstraction of a single cell in the sheet.

       CHILDREN NODES SHOULD HAVE CONTENTS SET TO EMPTY STRING.
    """
    def __init__(self, workbook: Workbook, sheet_name, cell_location, contents:str=''):
        self.workbook = workbook
        self.dependencies = set() #The cells this cell relies on
        self.cell_location = cell_location
        self.sheet_name = sheet_name
        self.content_str = contents
        workbook.set_cell_contents(sheet_name, cell_location, contents)

    def add_dependency(self, sheet_name, cell_location):
        """ Adds a dependency to this cell. Adds +1 to the first dependency.
            i.e. if A1 depends on B1, then A1's contents are =1+B1.
        """
        self.dependencies.add((sheet_name, cell_location))
        if self.content_str != '':
            self.content_str = self.content_str + f" + {sheet_name}!{cell_location}"
        else:
            self.content_str = f"=1 + {sheet_name}!{cell_location}"
        self.workbook.set_cell_contents(self.sheet_name, self.cell_location, self.content_str)
    
    def set_contents(self, contents: str):
        #Erases dependencies and allows user to set the contents directly. 
        #DONT MANUALLY ADD DEPENDENCIES HERE
        self.dependencies = []
        self.workbook.set_cell_contents(self.sheet_name, self.cell_location, contents)
        self.content_str = contents
    
    def get_value(self):
        return self.workbook.get_cell_value(self.sheet_name, self.cell_location)
        
class PerformanceTest(unittest.TestCase):
    def setUp(self):
        self.wb = Workbook()
         
    def test_cell_calculation_linear(self):
        #Test cell calculation timing through a long, linear chain

        #Create the root of the chain
        (_, sheet_name) = self.wb.new_sheet("Sheet1")
        root = PerformanceNode(self.wb, sheet_name, 'A1', '1')
        #Build the chain
        end_node = None
        chain_length = 1000
        for i in range(2, chain_length):
            curr_node = PerformanceNode(self.wb, sheet_name, f'A{i}')
            curr_node.add_dependency(sheet_name, f'A{i-1}')
            end_node = curr_node

        pr = cProfile.Profile()
        pr.enable()
        root.set_contents('2')
        pr.disable()
       
        self.assertEqual(end_node.get_value(), chain_length) #Assert to make sure the propagation happened.
       
        s = io.StringIO()
        sortby = 'cumulative'
        ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
        ps.print_stats(15)
        print(s.getvalue())
 

    def test_cell_calculation_bushy(self):
        #Test cell calculation timing through a short, bushy tree. (one change causes many updates)
        (_, sheet_name) = self.wb.new_sheet("Sheet1")
        root = PerformanceNode(self.wb, sheet_name, 'A1', '1')
        #Build the chain
        end_node = None
        parent_nodes = [root] # popping nodes out of to put children into
        
        height = 3 # can adjust for how tall you want the tree to be
        num_children = 3 # how many children each parent will have
        cell_counter = 5 # child cell location (i.e. A5, A6, .. A_)
        for i in range(height-1):
            temp = []
            while len(parent_nodes) != 0:
                curr_parent = parent_nodes.pop()
                for j in range(num_children):
                    new_child = PerformanceNode(self.wb, sheet_name, f'A{cell_counter}', '')
                    cell_counter += 1
                    new_child.add_dependency(curr_parent.sheet_name, curr_parent.cell_location)
                    temp.append(new_child)
            parent_nodes = temp
        
        pr = cProfile.Profile()
        pr.enable()
        root.set_contents('2')
        pr.disable()
        
        self.assertEqual(self.wb.get_cell_value(sheet_name, f'A{cell_counter-1}'), 2+(height-1)) #Assert the last leaf.

        s = io.StringIO()
        sortby = 'cumulative'
        ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
        ps.print_stats(15)
        print(s.getvalue())    

    def test_cycle_detection_linear(self):
        #Test cycle detection timing in a long, linear chain

        #Create the root of the chain
        (_, sheet_name) = self.wb.new_sheet("Sheet1")
        root = PerformanceNode(self.wb, sheet_name, 'A1', '')
        #Build the chain
        chain_length = 1000
        end_node = None
        for i in range(2, chain_length):
            curr_node = PerformanceNode(self.wb, sheet_name, f'A{i}')
            curr_node.add_dependency(sheet_name, f'A{i-1}')
            end_node = curr_node
        
        #Time how long detecting the cycle takes
        pr = cProfile.Profile()
        pr.enable()
        root.add_dependency(end_node.sheet_name, end_node.cell_location)
        pr.disable()
        
        cell_error_object = end_node.get_value()
        self.assertTrue(isinstance(cell_error_object, CellError))
        self.assertTrue(cell_error_object.get_error_type_string() == "#CIRCREF!")

        s = io.StringIO()
        sortby = 'cumulative'
        ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
        ps.print_stats(15)
        print(s.getvalue())


    def test_cycle_detection_bushy(self):
        #Test cycle detection timing in a short, bushy tree
        (_, sheet_name) = self.wb.new_sheet("Sheet1")
        root = PerformanceNode(self.wb, sheet_name, 'A1', '')
        #Build the chain
        parent_nodes = [root] # popping nodes out of to put children into
        
        height = 3 # can adjust for how tall you want the tree to be
        num_children = 3 # how many children each parent will have
        cell_counter = 5 # child cell location (i.e. A5, A6, .. A_)
        for i in range(height-1):
            temp = []
            while len(parent_nodes) != 0:
                curr_parent = parent_nodes.pop()
                for j in range(num_children):
                    new_child = PerformanceNode(self.wb, sheet_name, f'A{cell_counter}', '')
                    cell_counter += 1
                    new_child.add_dependency(curr_parent.sheet_name, curr_parent.cell_location)
                    temp.append(new_child)
            parent_nodes = temp
        
        #Create cycle by adding dependency from parent to leaf
        pr = cProfile.Profile()
        pr.enable()
        root.add_dependency(sheet_name, f'A{cell_counter-1}')
        pr.disable()

        cell_error_object = root.get_value()
        self.assertTrue(isinstance(cell_error_object, CellError))
        self.assertTrue(cell_error_object.get_error_type_string() == "#CIRCREF!")

        # Assert that the update time is within an acceptable threshold
        s = io.StringIO()
        sortby = 'cumulative'
        ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
        ps.print_stats(15)
        print(s.getvalue())    


    def test_cycle_detection_multiple_sheet_linear(self):
        # Tests when two sheets with linear dependencies are connected to form a cycle. 
        (_, sheet_name1) = self.wb.new_sheet("Sheet1")
        (_, sheet_name2) = self.wb.new_sheet("Sheet2")

        # Create linear dependencies in Sheet1
        root1 = PerformanceNode(self.wb, sheet_name1, 'A1', '')
        end_node1 = None
        for i in range(2, 100):
            curr_node = PerformanceNode(self.wb, sheet_name1, f'A{i}')
            curr_node.add_dependency(sheet_name1, f'A{i-1}')
            end_node1 = curr_node

        # Create linear dependencies in Sheet2
        root2 = PerformanceNode(self.wb, sheet_name2, 'B1', '')
        end_node2 = None
        for i in range(2, 100):
            curr_node = PerformanceNode(self.wb, sheet_name2, f'B{i}')
            curr_node.add_dependency(sheet_name2, f'B{i-1}')
            end_node2 = curr_node

        # Interlink the last nodes of each sheet to the first node of the other sheet and time cycle detection
        root1.add_dependency(sheet_name2, end_node2.cell_location)
        cycle_detection_time = timeit.timeit(
            lambda: root2.add_dependency(sheet_name1, end_node1.cell_location),
            number=1
        )
        pr = cProfile.Profile()
        pr.enable()
        root2.add_dependency(sheet_name1, end_node1.cell_location)
        pr.disable()

        # Check for cycle errors
        cycle_error1 = end_node1.get_value()
        cycle_error2 = end_node2.get_value()

        self.assertTrue(isinstance(cycle_error1, CellError))
        self.assertTrue(cycle_error1.get_error_type_string() == "#CIRCREF!")
        self.assertTrue(isinstance(cycle_error2, CellError))
        self.assertTrue(cycle_error2.get_error_type_string() == "#CIRCREF!")

        s = io.StringIO()
        sortby = 'cumulative'
        ps = pstats.Stats(pr, stream=s).sort_stats(sortby)
        ps.print_stats(15)
        print(s.getvalue())
    
if __name__ == '__main__':
    unittest.main()