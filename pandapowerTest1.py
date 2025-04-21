import pandapower as pp
import pandapower.shortcircuit as sc

import pandas as pd
import numpy as np


#add the full dataframe
df = pd.read_excel('Load List.xlsx', sheet_name="Buses")

# Filter the full dataframe with only busses
bus_df = df[df['Type']=='Bus']


# Create empty network
net = pp.create_empty_network()

# Creating the busses from the dataframe
for index, row in bus_df.iterrows():
    pp.create_bus(net, name=row['Description'], vn_kv=row['Voltage']*.001, type="b")
    
print (net.bus)
