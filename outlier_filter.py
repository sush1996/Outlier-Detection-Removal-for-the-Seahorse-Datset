"""
Author : Sushruth R P (Electrical and Systems Engineering Department, University of Pennsylavania)

The code below is to filter the OCR and ECAR data of all the outlier values.
The entire well is discarded both from OCR and ECAR data even if a single value in the well is an outlier
Outliers are caclulated by checking if they exceed the lower and the upper bounds of the group
Upper bound is calculated as the sum of the maximum value in the group and 1.5* IQR for that group
Upper bound is calculated as the difference of the minimum value in the group and 1.5* IQR for that group  

The resulting output file is stored in the same location as this python file. Look for the file "Outliers.xlsx" 
"""

import pandas as pd
import numpy as np
from scipy.stats import iqr

def Q_1(df_values):
    return df_values.quantile(0.25)

def Q_3(df_values):
    return df_values.quantile(0.75)

def excel_col_ind():
    
    i = raw_input("Enter the starting excel column numbers in the format (C AS ... etc):") #C AS CI DY FO HE 
    j = raw_input("Enter the terminating excel column numbers in the format: (AR CH ... etc):") #AR CH DX FN HD IT

    i = list(i.split(" "))
    j = list(j.split(" "))

    return i, j

def ocr_ecar(filepath):
	
	i, j = excel_col_ind()
		
	df_ocr = pd.DataFrame()
	df_ecar = pd.DataFrame()
	df_iqr = pd.DataFrame()

	params_ocr = []
	params_ecar = []
	param_keys = []

	ocr_lb_df = pd.DataFrame()
	ocr_lb_df_ex = pd.DataFrame()
	ocr_ub_df = pd.DataFrame()
	ocr_ub_df_ex = pd.DataFrame()

	ecar_lb_df = pd.DataFrame()
	ecar_lb_df_ex = pd.DataFrame()
	ecar_ub_df = pd.DataFrame()
	ecar_ub_df_ex = pd.DataFrame()

	ocr_ex = pd.DataFrame()
	ecar_ex = pd.DataFrame()

	for iter_i, iter_j in zip(i, j):

		group_columns = iter_i+":"+iter_j
		group_columns = 'A,'+ group_columns #including column A in the analysis for fetching OCR and ECAR data

		data = pd.read_excel(filepath, sheet_name = "Combined", skiprows = 1, usecols = group_columns)
		col_names = data.keys()

		ocr =  data[data['Normalized to 3rd'] == '%OCR to baseline']		#Fetching OCR data 
		ecar =  data[data['Normalized to 3rd'] == '%ECAR to baseline']		#Fetching ECAR data
	        
		ocr = ocr.iloc[:,1:]		#Skipping the first column
		ecar = ecar.iloc[:,1:]		#Skipping the first column

		
		#ocr_ex and ecar_ex contains all the tables of data as specified by the excel columnn indexes 
		
		ocr_ex = pd.concat([ocr_ex, ocr], axis =  1) 	#Concatenating tables of OCR data
		ecar_ex = pd.concat([ecar_ex, ecar], axis = 1)	#Concatenating tables of ECAR data

		#ocr_iqr and ecar_iqr have IQR values for the OCR and ECAR data
		ocr_iqr = ocr.apply(iqr, axis = 1).values
		ecar_iqr = ecar.apply(iqr, axis = 1).values
	        
		#<ocr, ecar>_q1 and <ocr, ecar>_q3 contains the first and the third quartiles of the ocr and ecar data respectively
		ocr_q1 = ocr.apply(Q_1, axis = 1).values
		ocr_q3 = ocr.apply(Q_3, axis = 1).values
		ecar_q1 = ecar.apply(Q_1, axis = 1).values
		ecar_q3 = ecar.apply(Q_3, axis = 1).values

		#<ocr, ecar>_lb and <ocr, ecar>_ub contains the lower and the upper bounds of the data 
		ocr_lb = ocr_q1 - 1.5*ocr_iqr
		ocr_ub = ocr_q3 + 1.5*ocr_iqr
		ecar_lb = ecar_q1 - 1.5*ecar_iqr
		ecar_ub = ecar_q3 + 1.5*ecar_iqr

		# Repeating and concatenating lower and upper bounds of ocr and ecar for element-wise comparison with the actual ocr and ecar data
		ocr_lb_df = pd.concat([pd.DataFrame(ocr_lb)]*len(ocr.keys()), axis = 1)
		ocr_lb_df_ex = pd.concat([ocr_lb_df_ex, ocr_lb_df], axis = 1)
	 	ocr_ub_df = pd.concat([pd.DataFrame(ocr_ub)]*len(ocr.keys()), axis = 1)
	 	ocr_ub_df_ex = pd.concat([ocr_ub_df_ex, ocr_ub_df], axis = 1)

	 	ecar_lb_df = pd.concat([pd.DataFrame(ecar_lb)]*len(ecar.keys()), axis = 1)
		ecar_lb_df_ex = pd.concat([ecar_lb_df_ex, ecar_lb_df], axis = 1)
	 	ecar_ub_df = pd.concat([pd.DataFrame(ecar_ub)]*len(ecar.keys()), axis = 1)
	 	ecar_ub_df_ex = pd.concat([ecar_ub_df_ex, ecar_ub_df], axis = 1)

	#Maintaining conssistent column keys and row indexes
	ocr_lb_df_ex.columns = ocr_ex.columns
	ocr_lb_df_ex.index = ocr_ex.index
	ocr_ub_df_ex.columns = ocr_ex.columns
	ocr_ub_df_ex.index = ocr_ex.index

	ecar_lb_df_ex.columns = ecar_ex.columns
	ecar_lb_df_ex.index = ecar_ex.index
	ecar_ub_df_ex.columns = ecar_ex.columns
	ecar_ub_df_ex.index = ecar_ex.index

	#Obtaining the boolean values of the OCR and ECAR after comparison to check for outliers
	# if True: it is an outlier, if False: not an outlier
	lt_ocr = ocr_ex < ocr_lb_df_ex 
	gt_ocr = ocr_ex > ocr_ub_df_ex
	bool_ocr = lt_ocr | gt_ocr

	lt_ecar = ecar_ex < ecar_lb_df_ex 
	gt_ecar = ecar_ex > ecar_ub_df_ex
	bool_ecar = lt_ecar | gt_ecar


	#Dropping the columns if and where even a single value in the well violates bounds
	#Keep columns of data where the result is False in all the rows of the column
	filter_ocr = bool_ocr == False
	ocr_filt = ocr_ex.where(filter_ocr)
	ocr_filt = ocr_filt.dropna(axis = 'columns')
	
	filter_ecar = bool_ecar == False
	ecar_filt = ecar_ex.where(filter_ecar)
	ecar_filt = ecar_filt.dropna(axis = 'columns')

	temp = pd.DataFrame(columns = ocr_ex.columns)
	temp = pd.concat([temp, ocr_filt], axis = 0, ignore_index = True, sort = False)

	#Concatenating the filtered OCR and ECAR data to be entered into the excel file 
	filt_data = pd.DataFrame(temp)
	filt_data = pd.concat([ocr_filt, ecar_filt], axis = 0, sort = False, keys= ['OCR', 'ECAR'])
	filt_data = filt_data.dropna(axis = 'columns')
	#ocr_ecar_filt = pd.concat([ocr_filt, ecar_filt], axis = 0, keys = ['OCR', 'ECAR'], columns = ocr_ex.keys(), sort = False)
	filt_data.to_excel('Outliers_test.xlsx')

def main():
	filepath = raw_input("Enter the file path:") #"/home/sushruth/Documents/Redox_Lab/Combined.xlsx"
	ocr_ecar(filepath)

if __name__ == "__main__":
	main()

