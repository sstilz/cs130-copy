import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import unittest
from sheets.graph import Graph
from sheets.cell import Cell

import unittest

class TestGraph(unittest.TestCase):
    def setUp(self):
        # Common setup code goes here
        self.graph = Graph()
    
    def tearDown(self):
        # Common teardown code goes here
        del self.graph


    def test_add_single_node(self):
        node = "A"
        self.graph.add_node(node)
        self.assertTrue(self.graph.is_in_graph(node))
        self.assertEqual(self.graph.get_children(node), [])
    
    def test_add_edge_two_nodes(self):
        u = "A"
        v = "B"
        self.graph.add_edge(u, v)
        self.assertTrue(self.graph.is_in_graph(u))
        self.assertTrue(self.graph.is_in_graph(v))
        self.assertEqual(self.graph.get_children(u), [v])
        self.assertEqual(self.graph.get_children(v), [])

    def test_clear_refs_criterion(self):
        # Test case 1: Delete nodes that are even
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 4)
        self.graph.add_edge(4, 5)

        crit_func = lambda x: x % 2 == 0
        self.graph.clear_refs_criterion(crit_func)
        self.assertEqual(self.graph.get_adj_list(), {1: [], 2:[3], 3: [], 4: [5], 5:[]})


    def test_tarjan_single_component(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)

        self.assertEqual(self.graph.get_sccs(), [{1, 2, 3}])

    def test_tarjan_disconnected_components(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)
        self.graph.add_edge(4, 5)

        self.assertCountEqual(self.graph.get_sccs(), [{1, 2, 3}, {4}, {5}])

    def test_tarjan_two_distinct_components(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)
        self.graph.add_edge(3, 4)
        self.graph.add_edge(4, 5)
        self.graph.add_edge(5, 4)

        self.assertCountEqual(self.graph.get_sccs(), [{4, 5}, {1, 2, 3}])
    
    def test_tarjan_multiple_cycles_in_cycle(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)
        self.graph.add_edge(3, 4)
        self.graph.add_edge(4, 1)

        self.assertCountEqual(self.graph.get_sccs(), [{1, 2, 3, 4}])

    def test_graph_cell(self):
        a = Cell("abc")
        b = Cell("abc")
        c = Cell("abc")

        self.graph.add_edge(a, b)
        self.graph.add_edge(b, c)
        self.graph.add_edge(c, a)

        self.assertCountEqual(self.graph.get_sccs(), [{a,b,c}])
    
    def test_get_scc_number_multiple_cycles_in_cycle(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)
        self.graph.add_edge(4, 5)
        self.graph.add_edge(5, 6)
        self.graph.add_edge(6, 4)

        
        # We determined these specific SCC numbers using the fact that SCC numbers are assigned in the order that a SCC is fully discovered.
        scc_number = self.graph.get_scc_number(1)
        self.assertEqual(scc_number, 0)

        scc_number = self.graph.get_scc_number(4)
        self.assertEqual(scc_number, 1)

        scc_number = self.graph.get_scc_number(6)
        self.assertEqual(scc_number, 1)
    
    def test_scc_dag(self):
        # create 2 distinct SCCs with an edge going from one to the other
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)
        self.graph.add_edge(4, 5)
        self.graph.add_edge(5, 6)
        self.graph.add_edge(6, 4)
        self.graph.add_edge(1, 4)    # edge going from one SCC to another

        self.assertCountEqual(self.graph.get_sccs(), [{1,2,3}, {4,5,6}])
        scc_123 = self.graph.get_scc_number(1)
        scc_456 = self.graph.get_scc_number(4)

        self.assertEqual(self.graph.get_scc_dag(), {scc_456: [], scc_123: [scc_456]})
    

    def test_topological_sort_straight_line_dependencies(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 4)

        self.assertEqual(self.graph.get_topo_sort(), [1, 2, 3, 4])
    
    def test_topological_sort_triangle_dependencies(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(1, 3)
        self.graph.add_edge(2, 3)

        self.assertEqual(self.graph.get_topo_sort(), [1, 2, 3])

    def test_topological_sort_multiple_dependencies(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(1, 3)
        self.graph.add_edge(2, 4)
        self.graph.add_edge(3, 4)


        self.assertEqual(self.graph.get_topo_sort(), [1, 2, 3, 4])

    # When given 2 disjoint SCCs, we have a DAG with no edges. Thus there is no "expected" output for topological sort.
    def test_topological_sort_graph_with_no_edges(self):
        self.graph.add_edge(1, 2)
        self.graph.add_edge(2, 3)
        self.graph.add_edge(3, 1)
        self.graph.add_edge(4, 5)
        self.graph.add_edge(5, 6)
        self.graph.add_edge(6, 4)

        self.assertEqual(len(self.graph.get_sccs()), 2)
        self.assertEqual(len(self.graph.get_topo_sort()), 6)

    def test_single_isolated_node_as_scc(self):
        # Test that a single isolated node is one SCC

        # Create a graph with a single isolated node
        node = "A"
        self.graph.add_node(node)

        # Check that there is exactly one SCC, and it contains the isolated node
        sccs = self.graph.get_sccs()
        self.assertEqual(len(sccs), 1)
        self.assertEqual(sccs[0], {node})

        # Check SCC number for the isolated node
        scc_number = self.graph.get_scc_number(node)
        self.assertEqual(scc_number, 0)

    def test_clear_refs(self):
        node = "A"
        self.graph.add_node(node)

        node2 = "B"
        self.graph.add_node(node2)

        self.graph.add_edge(node2, node)

        self.graph.clear_refs(node)
        self.assertNotIn(node, self.graph.get_children(node2))
        self.assertTrue(self.graph.is_in_graph(node))
        self.assertTrue(self.graph.is_in_graph(node2))

    def test_rename_cell_existing_cell(self):
        self.graph.add_node("A")
        self.graph.add_node("B")
        self.graph.add_edge("A", "B")

        self.graph.rename_cell("A", "C")
        self.assertNotIn("A", self.graph.get_all_nodes())
        self.assertIn("C", self.graph.get_all_nodes())
        self.assertIn("B", self.graph.get_children("C"))

    def test_rename_cell_non_existing_cell(self):
        result = self.graph.rename_cell("C", "D")
        self.assertIsNone(result)
        self.assertNotIn("D", self.graph.get_all_nodes())

    def test_rename_cell_complex(self):
        # Add nodes and edges to the graph
        self.graph.add_node("A")
        self.graph.add_node("B")
        self.graph.add_node("C")
        self.graph.add_edge("A", "B")
        self.graph.add_edge("B", "C")

        # Rename cell "B" to "D"
        self.graph.rename_cell("B", "D")

        # Check that the old cell name is not in the graph
        self.assertNotIn("B", self.graph.get_all_nodes())

        # Check that the new cell name is in the graph
        self.assertIn("D", self.graph.get_all_nodes())

        # Check that the edges are updated correctly
        self.assertIn("D", self.graph.get_children("A"))
        self.assertIn("C", self.graph.get_children("D"))
        
    def test_get_all_nodes(self):
        self.graph.add_node("A")
        self.graph.add_node("B")
        self.graph.add_edge("A", "B")

        nodes = self.graph.get_all_nodes()
        self.assertEqual(len(nodes), 2)
        self.assertIn("A", nodes)
        self.assertIn("B", nodes)
        
    def test_get_all_nodes(self):
        self.graph.add_node("A")
        self.graph.add_node("B")
        self.graph.add_edge("A", "B")

        nodes = self.graph.get_all_nodes()
        self.assertEqual(len(nodes), 2)
        self.assertIn("A", nodes)
        self.assertIn("B", nodes)
    
    def test_adjacency_list_copied_after_rename(self):
        self.graph.add_node("A")
        self.graph.add_node("B")
        self.graph.add_edge("A", "B")

        old_refs = self.graph.get_children("A")
        self.graph.rename_cell("A", "C")

        self.assertEqual(self.graph.get_children("C"), old_refs)
        self.assertEqual(self.graph.get_children("C"), ["B"])

    def test_rename_to_self(self):
        self.graph.add_node("A")
        self.graph.add_node("B")
        self.graph.add_edge("A", "B")
        # Changed so it doesn't raise an error (in case for example, we create a circular reference through renaming sheet)
        try:
            self.graph.rename_cell("A", "A")
        except:
            self.fail("rename_cell raised an exception when renaming to itself")

if __name__ == '__main__':
    unittest.main()
