import pandas as pd
from scipy.stats import iqr


def Q_1(df_values):
    return df_values.quantile(0.25)

def Q_3(df_values):
    return df_values.quantile(0.75)

def outlier_bounds(fd,fn):
    fp = fd + "/" + fn
    
    i = raw_input("Enter the starting excel column numbers in the format A AR ... etc:") 
    j = raw_input("Enter the terminating excel column numbers in the format: AS CH ... etc:") 

    i = list(i.split(" "))
    j = list(j.split(" "))
    
    df_ocr = pd.DataFrame()
    df_ecar = pd.DataFrame()
    df_iqr = pd.DataFrame()

    params_ocr = []
    params_ecar = []
    param_keys = []

    for iter_i, iter_j in zip(i, j):
        
        ps = iter_i+":"+iter_j
        ps = 'A,'+ps
        
        data = pd.read_excel(fn, sheet_name = "Combined", skiprows = 1, usecols = ps)
        
        ocr =  data[data['Normalized to 3rd'] == '%OCR to baseline']
        ecar =  data[data['Normalized to 3rd'] == '%ECAR to baseline']
        
        ocr = ocr.iloc[:,1:]
        ecar = ecar.iloc[:,1:]

        ocr_iqr = ocr.apply(iqr, axis = 1).values
        ecar_iqr = ecar.apply(iqr, axis = 1).values
        
        ocr_q1 = ocr.apply(Q_1, axis = 1).values
        ocr_q3 = ocr.apply(Q_3, axis = 1).values
        ecar_q1 = ecar.apply(Q_1, axis = 1).values
        ecar_q3 = ecar.apply(Q_3, axis = 1).values

        ocr_ub = ocr_q3 + 1.5*ocr_iqr
        ocr_lb = ocr_q1 - 1.5*ocr_iqr

        ecar_ub = ecar_q3 + 1.5*ecar_iqr
        ecar_lb = ecar_q1 - 1.5*ecar_iqr

        ocr_q1 = pd.Series(ocr_q1)
        ocr_q3 = pd.Series(ocr_q3)    
        ocr_iqr = pd.Series(ocr_iqr)
        ocr_lb = pd.Series(ocr_lb)
        ocr_ub = pd.Series(ocr_ub)

        ecar_q1 = pd.Series(ecar_q1)
        ecar_q3 = pd.Series(ecar_q3)
        ecar_iqr = pd.Series(ecar_iqr)
        ecar_lb = pd.Series(ecar_lb)
        ecar_ub = pd.Series(ecar_ub)
        
        params_ocr = params_ocr + [ocr_q1, ocr_q3, ocr_iqr, ocr_ub, ocr_lb]
        params_ecar = params_ecar + [ecar_q1, ecar_q3, ecar_iqr, ecar_ub, ecar_lb]
        param_keys = param_keys + ['Q1', 'Q3', 'IQR', 'LB', 'UB']

    df_ocr = pd.concat(params_ocr, axis = 1, keys = param_keys)
    df_ecar = pd.concat(params_ecar, axis = 1, keys = param_keys)
        
    df_iqr = pd.concat([df_ocr, df_ecar], axis = 0, keys = ['%OCR to baseline', '%ECAR to baseline'])
    df_iqr.to_excel("outliers.xlsx")

def main():
    
    filepath = raw_input("Enter the file directory of the excel file: ")
    filename = raw_input("Enter the name of the excel file: ")
    outlier_bounds(filepath, filename)
    print("Successful, check for the outliers file")

if __name__ == "__main__":
    main()