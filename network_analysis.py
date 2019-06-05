import xlrd
import networkx as nx
import matplotlib.pyplot as plt
import operator
import community

# Creating Empty Graph
G = nx.Graph()

# Definition of excel spreadsheets
wb_nodes = xlrd.open_workbook('D:/coding/network_analysis/code/nodes.xlsx')
wb_edges = xlrd.open_workbook('D:/coding/network_analysis/code/edges.xlsx')
sheet_nodes = wb_nodes.sheet_by_index(0)
sheet_edges = wb_edges.sheet_by_index(0)


# Extracting nodes into node list
nodes = []

for i in range(1, sheet_nodes.nrows):
    nodes.append(sheet_nodes.cell_value(i, 0))

# Extracting edges
for i in range(1, sheet_edges.nrows):
    G.add_edge(sheet_edges.cell_value(i, 0), sheet_edges.cell_value(i, 1))

# Adding nodes to a graph
G.add_nodes_from(nodes)
print(nx.info(G))

# Density from 0-1, 0 - not connected, 1 - perfectly connected network
density = nx.density(G)
print("Network density:", density)

# Shortest path
sh_pth_spoge_tortuga = nx.shortest_path(G, source="Spoge", target="Tortuga")
print("Shortest path between Spoge and Tortuga: ", sh_pth_spoge_tortuga)
print("length of that path:", len(sh_pth_spoge_tortuga)-1)

# Diameter - longest of all shortest paths - nx.diameter only works if whole graph is one component
print("Network diameter of largest component:", nx.diameter(G))

# Transitivity - ratio of triangles over all possible triangles
triadic_closure = nx.transitivity(G)
print("Triadic closure:", triadic_closure * 100, "%")

# Triangles in network for each node
triangles = nx.triangles(G)
print("Triangles in graph: ", nx.triangles(G))

# Clustering of a node
clustering = nx.clustering(G)
print("Clustering coefficient : ", clustering)

# Average clustering
avg_clustering = nx.average_clustering(G)
print("Average clustering of a graph G is:", avg_clustering)

# Efficiency
print("Efficiency of graph:", nx.global_efficiency(G))

# Degree - node's degree is the sum of its edges


def takesecond(elem):
    return elem[1]


degree_list = list(G.degree())
# node_sizes is list used for drawing graph, its calculated from list of degree's elements * 10
node_sizes = []
for i in degree_list:
    node_sizes.append(i[1]*10)

# sort and print out degree list
degree_list.sort(key=takesecond, reverse=True)
print("Nodes degree's: ", degree_list)


# CENTRALITY methods: #

# Closeness centrality
close_list = nx.closeness_centrality(G)
close_list = sorted(close_list.items(), key=operator.itemgetter(1), reverse=True)
print("Closeness centrality: ", close_list)

# Degree centrality
deg_list = nx.degree_centrality(G)
deg_list = sorted(deg_list.items(), key=operator.itemgetter(1), reverse=True)
print("Degree centrality: ", deg_list)

# Eigenvector - centrality looks at a combination of a node’s edges and the edges of that node’s neighbors.
eigen_list = (nx.eigenvector_centrality(G))
eigen_list = sorted(eigen_list.items(), key=operator.itemgetter(1), reverse=True)
print("Eigenvector centrality: ", eigen_list)

# betweenness centrality - betweenness centrality is a measure of centrality in a graph based on shortest paths.

betweenness_list = nx.betweenness_centrality(G)
betweenness_list = sorted(betweenness_list.items(), key=operator.itemgetter(1), reverse=True)
print("Betweenness centrality: ", betweenness_list)


# COMMUNITY #

# community
part = community.best_partition(G)
color_list = list(part.values())

# Modularity
mod = community.modularity(part,G)
print("modularity:", mod)


# Draw graph
nx.draw_networkx(
    G,
    edge_color='gray',
    node_size=node_sizes,
    font_size=6,
    font_color='black',
    node_color=color_list
)

plt.show()