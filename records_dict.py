#!/usr/bin/env python
# coding: utf-8

# In[3]:


f_p = 'C:/Users/SAMUEL N-AMOABENG/anaconda3/envs/admission-dec/Yue - Admit Project/China Example'


# In[2]:


import os
import pandas as pd

def extract_data_from_files(directory):
    # Initialize an empty dictionary to store the data
    records = {}
    
    # Loop through each file in the directory
    for file_name in os.listdir(directory):
        # Construct the file path
        file_path = os.path.join(directory, file_name)
        
        # Read the Excel file into two DataFrames with headers 0 and 1
        r1 = pd.read_excel(file_path, header=0)
        r2 = pd.read_excel(file_path, header=1)
        
        # Extract reviewer 1 name from file name
        reviewer1 = file_name.split('_')[1]
        
        # Check if the DataFrames contain the required columns
        if 'Rating' in r1.columns and 'Name'and 'Reader 2 Name' and 'Institution 1 GPA (4.0 Scale)' in r2.columns:
            # Extract the rating and name data from the DataFrames
            rating = r1['Rating'][1:].to_numpy()
            name = r2['Name'].to_numpy()
            reader2 = r2['Reader 2 Name'].to_numpy()
            gpa = r2['Institution 1 GPA (4.0 Scale)'].to_numpy()
            
            # Create a dictionary for the file data
            file_data = {}
            for i in range(len(name)):
                file_data[name[i]] = {'Reviewer1': reviewer1,
                                      'Rating': rating[i],
                                      'Reader2': reader2[i],
                                      'GPA': gpa[i]}
            
            # Add the file data to the main data dictionary
            records[file_name] = file_data
    
    # Return the main data dictionary
    return records


# In[4]:


extract_data_from_files(f_p)


# In[ ]:





# In[ ]:




