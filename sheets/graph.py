class Graph:
    def __init__(self):
        self.graph = dict()
        self.sccs = []  # List to store SCCs

        self.node_to_scc_num = dict()
        self.scc_dag = dict()
        self.topo_sort = list()

        self.updated = True
    
    def rename_cell(self, old_cell, new_cell):
        if old_cell not in self.graph:
            return None
        
        if new_cell not in self.graph:
            self.graph[new_cell] = list(self.graph[old_cell])

        del self.graph[old_cell]

        for node in self.graph:
            if old_cell in self.graph[node]:
                self.graph[node].remove(old_cell)
                self.graph[node].append(new_cell)

        self.updated = False

    def get_all_nodes(self):
        return list(self.graph.keys())
    
    def add_node(self, node):
        if node in self.graph:
            return None
        self.graph[node] = list()

        self.updated = False

    def add_edge(self, u, v):
        """
        Adds edge u ---> v
        """
        self.add_node(u)
        self.add_node(v)
            
        self.graph[u].append(v)

        self.updated = False

    def clear_refs(self, node):
        # clears all edges going into node
        for other_node in self.graph:
            if node in self.graph[other_node]:
                self.graph[other_node].remove(node)
        
        self.updated = False

    def clear_refs_criterion(self, crit_func):
        to_clear = list()
        for node in self.graph:
            if crit_func(node):
                to_clear.append(node)

        for node in to_clear:
            self.clear_refs(node)

    def tarjan(self):
        index = 0
        stack = []
        low_link = {}
        on_stack = {}
        result = []
        curr_index = {}

        call_stack = []
        for node in self.graph:
            if low_link.get(node, -1) == -1:
                call_stack.append((node, 0))

                while call_stack:
                    u, neighbor_index = call_stack.pop()

                    # First time visiting u
                    if neighbor_index == 0:
                        index += 1
                        curr_index[u] = index
                        low_link[u] = index
                        stack.append(u)
                        on_stack[u] = True
                    
                    # We just returned from a neighbor
                    if neighbor_index > 0:
                        previous_neighbor = self.graph[u][neighbor_index - 1]
                        if on_stack[previous_neighbor]:
                            low_link[u] = min(low_link[u], low_link[previous_neighbor])

                    # Skip all neighbors that have already been visited
                    while neighbor_index < len(self.graph[u]):
                        v = self.graph[u][neighbor_index]
                        if low_link.get(v, -1) == -1:
                            break
                        if on_stack[v]:
                            low_link[u] = min(low_link[u], low_link[v])
                        neighbor_index += 1

                    # If we still have neighbors to visit
                    if neighbor_index < len(self.graph[u]):
                        v = self.graph[u][neighbor_index]
                        call_stack.append((u, neighbor_index + 1))
                        call_stack.append((v, 0))
                        continue
                    
                    # No more neighbors to visit. We can check if we have an SCC
                    if low_link[u] == curr_index[u]:
                        scc = []
                        while True:
                            node = stack.pop()
                            on_stack[node] = False
                            scc.append(node)

                            self.node_to_scc_num[node] = len(result)

                            if node == u:
                                break
                            
                        result.append(set(scc))

        return result
    def build_scc_dag(self, sccs):
        scc_dag = dict()
        for scc_num, scc in enumerate(sccs):
            scc_dag[scc_num] = set()
            for node in scc:
                for neighbor in self.graph[node]:
                    if neighbor not in scc:
                        scc_dag[scc_num].add(self.node_to_scc_num.get(neighbor))
        
        return scc_dag

    def topological_sort(self, scc_dag):
        in_degree = {node: 0 for node in scc_dag}
        for node in scc_dag:
            for neighbor in scc_dag[node]:
                in_degree.setdefault(neighbor, 0)

        for node in scc_dag:
            for neighbor in scc_dag[node]:
                in_degree[neighbor] += 1

        queue = [node for node in in_degree if in_degree[node] == 0]

        sorted_order = []
        while queue:
            node = queue.pop(0)
            sorted_order.append(node)

            for neighbor in scc_dag.get(node, []):
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)

        return sorted_order
    
    def get_node_topo_from_scc_topo(self, sccs, scc_topo_sort):
        ans = list()
        for scc_num in scc_topo_sort:
            for node in sccs[scc_num]:
                ans.append(node)
        return ans
    
    def update(self):
        self.sccs = self.tarjan()
        self.scc_dag = self.build_scc_dag(self.sccs)
        scc_topo_sort = self.topological_sort(self.scc_dag)

        self.topo_sort = self.get_node_topo_from_scc_topo(self.sccs, scc_topo_sort)

        self.updated = True

    def get_sccs(self):
        if not self.updated:
            self.update()
        return self.sccs

    def get_scc_number(self, u):
        if not self.updated:
            self.update()

        return self.node_to_scc_num.get(u)

    def get_topo_sort(self):
        if not self.updated:
            self.update()

        return self.topo_sort
    
    def get_component(self, node):
        if not self.updated:
            self.update()

        if self.is_in_graph(node):
            return self.sccs[self.node_to_scc_num.get(node)]
        
    def is_in_graph(self, node):
        return node in self.graph
    
    def get_children(self, node):
        return list(self.graph.get(node, []))

    def get_adj_list(self):
        return { node : list(children) for node, children in self.graph.items() }
    
    def get_scc_dag(self):
        if not self.updated:
            self.update()
        return { scc_num : list(children) for scc_num, children in self.scc_dag.items()}