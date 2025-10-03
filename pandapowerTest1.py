import pandapower as pp
import pandapower.shortcircuit as sc

import pandas as pd
import numpy as np

# Create empty network
net = pp.create_empty_network()


# Load the data from your Excel file
# Ensure your Excel file is named 'cLoad List.xlsx' and the sheet is 'Cables'
cable_df = pd.read_excel("Load List.xlsx", sheet_name="cable_lib", index_col=0)

# 2. Iterate through the DataFrame and create a standard type for each cable
for cable_name, row in cable_df.iterrows():
    # 'line' indicates that this is a type for a pandapower line element
    pp.create_std_type(
        net=net,  # Pass None to add it to the library
        name=cable_name,
        element="line",
        data={
            "r_ohm_per_km": row["r_ohm_per_km"],
            "x_ohm_per_km": row["x_ohm_per_km"],
            "c_nf_per_km": row["c_nf_per_km"],
            "max_i_ka": row["max_i_ka"],
            "type": "cs",  # "cs" for cable system (cable)
        }
    )

#add the full dataframe
df = pd.read_excel('Load List.xlsx', sheet_name="Buses")

# Filter the full dataframe with only busses
bus_df = df[df['Type']=='Bus']

# Creating the busses from the dataframe
for index, row in bus_df.iterrows():
    pp.create_bus(net, name=row['Description'], vn_kv=row['Voltage']*.001, type="b")

# Use one of the custom cable names defined in your Excel
pp.create_line(net, from_bus=1, to_bus=2, length_km=5.0, std_type="CustomCable1")
pp.create_line(net, from_bus=0, to_bus=1, length_km=5.0, std_type="CustomCable1")

# Display the network to verify
print (net.bus)

'pp.runpp(net)'