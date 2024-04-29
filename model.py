# Import Packages
import matplotlib.pyplot as plt
import gurobipy as gp
from gurobipy import GRB
import pandas as pd

# Read Data
item_info = pd.read_excel('Aggregated Supply Information.xlsx')
item_info['Volume (Cubic In.)'] = item_info['Volume (Cubic In.)'].fillna(80)
item_info.dropna(axis=0, subset=['Par', 'Total Picks'], inplace=True)
item_info.reset_index(inplace=True)

items = [i for i in item_info['Item Number']]
sizes = {i: s for i, s in zip(item_info['Item Number'], item_info['Volume (Cubic In.)'])}
quantities = {i: q for i, q in zip(item_info['Item Number'], item_info['Par'])}
frequencies = {i: f for i, f in zip(item_info['Item Number'], item_info['Total Picks'])}

location_info = pd.read_excel('Current Layout Information.xlsx')

shelves = [s for s in location_info['Shelf Number']]
distances = {s: d for s, d in zip(location_info['Shelf Number'], location_info['Distance'])}
capacities = {s: c for s, c in zip(location_info['Shelf Number'], location_info['Scaled Capacity'])}

# Create Model
m = gp.Model(name="Supply Locations")

# Define Decision Variables
l_list = [(i,s) for i in items for s in shelves]
L = m.addVars(l_list, name='L', vtype=GRB.BINARY)

# Define Objective
m.setObjective(gp.quicksum(L[i,s]*frequencies[i]*distances[s] for i in items for s in shelves), GRB.MINIMIZE)

# Define Constraints
for i in items:
    m.addConstr(gp.quicksum(L[i,s] for s in shelves) == 1)
for s in shelves:
    m.addConstr(gp.quicksum(L[i,s]*sizes[i]*quantities[i] for i in items) <= capacities[s])
    m.addConstr(gp.quicksum(L[i,s] for i in items) <= 50)

# Run Model
m.optimize()
print("Optimal solution:")
for v in m.getVars():
    if v.x == 1:
        print("%s = %g" % (v.varName, v.x))
#print(m.printStats())
print("Optimal objective value:\n{}".format(m.objVal))

# Extract optimal solution
optimal_solution = {(i, s): L[i, s].x for i in items for s in shelves}
filtered_pairs = [(i, s) for (i, s), value in optimal_solution.items() if value > 0.5]

# Output to Excel
output_df = pd.DataFrame(filtered_pairs, columns=['Item', 'Shelf'])
output_df.to_excel("Optimization Output.xlsx", index=False)

# Initialize dictionaries to store item counts and utilized capacity per shelf
shelf_item_counts = {s: 0 for s in shelves}
shelf_utilized_capacity = {s: 0 for s in shelves}

# Calculate item counts and utilized capacity per shelf
for (i, s) in filtered_pairs:
    shelf_item_counts[s] += 1
    shelf_utilized_capacity[s] += sizes[i] * quantities[i]

# Sort shelves based on distances
sorted_shelves = [s for s,_ in sorted(distances.items(), key=lambda x: x[1])]

# Reorder the dictionaries based on sorted shelves
sorted_shelf_utilized_capacity = {s: shelf_utilized_capacity[s]/capacities[s] for s in sorted_shelves}

# Visualize item counts per shelf with sorted shelves
plt.figure(figsize=(10, 6))
plt.bar(range(len(sorted_shelves)), shelf_item_counts.values(), color='skyblue')
plt.xlabel('Shelf Number')
plt.ylabel('Total Quantity of Items')
plt.title('Distribution of Items Across Shelves (Sorted by Distance)')
plt.xticks(range(len(sorted_shelves)), sorted_shelves)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.show()

# Visualize utilized capacity per shelf with sorted shelves
plt.figure(figsize=(10, 6))
plt.bar(range(len(sorted_shelves)), sorted_shelf_utilized_capacity.values(), color='lightgreen')
plt.xlabel('Shelf Number')
plt.ylabel('Utilized Capacity')
plt.title('Utilized Capacity of Each Shelf (Sorted by Distance)')
plt.xticks(range(len(sorted_shelves)), sorted_shelves)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.show()
