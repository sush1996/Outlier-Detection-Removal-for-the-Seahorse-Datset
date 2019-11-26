import pandas as pd
import numpy as np
from collections import OrderedDict
from scipy.stats import iqr

#Function to return an ordered set
def ordered_set(df_values):
	return list(OrderedDict.fromkeys(df_values))

#Function to reorganize the given excel file
def reorganize_excel(filepath):
    xls = pd.ExcelFile(filepath)
    xl_sh_0 = pd.read_excel(xls, sheet_name = 0, skiprows = 5)
    df = pd.DataFrame(xl_sh_0)
    
    group_names = ordered_set(df['Group Name'].values)
    measurements = len(ordered_set(df['Measurement'].values))
    measurement = ordered_set(df['Measurement'].values)
    time_stamps = ordered_set(df['Time'].values)
    ocr_group_collection = OrderedDict()
    ecar_group_collection = OrderedDict()
    
    for i in group_names:
    	group_df = df.loc[df['Group Name'] == i]
    	well = ordered_set(group_df['Well'].values)
    	ocr_temp_df = pd.DataFrame(index = range(1, measurements+1) ,columns = well)   #Creating a temporary dataframe for OCR
    	ecar_temp_df = pd.DataFrame(index = range(1, measurements+1) ,columns = well)  #Creating a temporary dataframe for ECAR
    
    	for num, j in enumerate(well):
            print(num, j)
            well_df = group_df.loc[group_df['Well'] == j]
    		
            ocr_val = well_df['OCR'].values
            ocr_temp_df[well[num]] = ocr_val
    		
            ecar_val = well_df['ECAR'].values
            ecar_temp_df[well[num]] = ecar_val

    	ocr_group_collection[i] = ocr_temp_df         #Creating a dictionary which has OCR data for all groups
    	ecar_group_collection[i] = ecar_temp_df       #Creating a dictionary which has ECAR data for all groups
    	
        '''
    	ocr_temp_df.columns = i + ", " + ocr_temp_df.columns
    	ecar_temp_df.columns = i + ", " + ecar_temp_df.columns
        '''
    print(ocr_group_collection)
    #measurement = pd.Series(measurement)
    time_stamps = pd.Series(time_stamps)
    
    # Fetching OCR Data
    ocr = pd.DataFrame() 
    ocr.insert(0, "Time", time_stamps)
    ocr.index+=1

    for i in range(0, len(ocr_group_collection)):
    	ocr = pd.concat([ocr,ocr_group_collection.values()[i]], axis = 1)
    '''
    b = ocr.columns.str.split(', ', expand=True).values
    ocr.columns = pd.MultiIndex.from_tuples([('', x[0]) if pd.isnull(x[1]) else x for x in b])
    '''
    ocr.index = np.arange(1, len(ocr)+1)
    ocr.index.name = 'Measurement'
    # Fetchng ECAR Data
    ecar = pd.DataFrame()
    ecar.insert(0, "Time", time_stamps)
    ecar.index+=1
    for i in range(0, len(ecar_group_collection)):
    	ecar = pd.concat([ecar,ecar_group_collection.values()[i]], axis = 1)
    '''
    b = ecar.columns.str.split(', ', expand=True).values
    ecar.columns = pd.MultiIndex.from_tuples([('', x[0]) if pd.isnull(x[1]) else x for x in b])
    '''
    ecar.index = np.arange(1, len(ecar)+1)
    ecar.index.name = 'Measurement'
    #Enter OCR and ECAR data into Excel
    xls = pd.concat([ocr, ecar], keys = ["OCR", "ECAR"])
    xls.to_excel("reformatted_wave.xlsx")

def main():
    
    filepath = raw_input("Enter the filepath of the excel file to be reformatted:")
    reorganize_excel(filepath)
    print("Reformat successful, check for the file in the filepath")

if __name__ == "__main__":
    main()
