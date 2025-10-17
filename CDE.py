import pandas as pd
import numpy as np

def split_point_names(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file, engine='openpyxl')
    # Add the index column
    df.reset_index(inplace=True)
    df.to_excel(r'C:\GHUFRAN\Old\PythonScripting\Niagara\Evap_Cooler_Points_index.xlsx')
    list1=[]
    for index, row in df.iterrows():
        if (df.loc[index, 'Point_Name']) == 'DAHU':
            list1.append(index)
            print(list1)

    df_null = df[df['Point_Name'].isnull()]
    list1.append(df_null.index.tolist())
    print(list1)
    df_sliced= df.loc[list1[0]:list1[1][0]-1]
    df_sliced.to_excel(r'C:\GHUFRAN\Old\PythonScripting\Niagara\Evap_Cooler_Points_sliced.xlsx') 
    list_columns = df_sliced.columns.tolist()
    print('length of columns:', len(list_columns))   
    # Split the 'Point_Name' column using regex
    split_df = df_sliced['Point_Name'].str.split(r'\[(.*?)\]', expand=True)
    # Convert to object type and replace NaN with None
    split_df = split_df.astype(object).mask(lambda x: x.isna(), None)
    #Assigning a newws for  column with first split part
    split_df_1=split_df.assign(new_col=split_df[0].astype(str))
    # All the rows where new_col is 'None' are replaced wiyth the value in the first row
    for index, rows in split_df_1.iterrows():
        split_df_1.loc[index,'new_col'] = split_df_1.loc[list1[0],'new_col'] 

    # Replace None with empty string
    split_df_1 = split_df_1.fillna('')
    
    # Generate a list of column names based on the number of splits
    list_column_names = [i for i in range(split_df_1.shape[1]-1)]  # Exclude 'new_col'

    list_val =list_column_names
    
    #list_val =[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10] # Is this hard-coded??
    list1=[]
    list2=[]
    for index,rows in split_df_1.iterrows():
        key = split_df_1.loc[index,'new_col']
        for x in list_val:
            list1.append(split_df_1.loc[index,x])
        value =list1
        dicT= {key : value} 
        list1=[]
        list2.append(dicT)
    print(list2)

    

    # Save to a new Excel file
    split_df_1.to_excel(output_file, index=False)
    print(f"Split data saved to {output_file}")

    
    


# Example usage:
#split_point_names('C:\GHUFRAN\Old\PythonScripting\Niagara\Evap_Cooler_Points.xlsx', 'C:\GHUFRAN\Old\PythonScripting\Niagara\Evap_Cooler_Points_1.xlsx')

split_point_names(
    r'C:\GHUFRAN\Old\PythonScripting\Niagara\Evap_Cooler_Points.xlsx',
    r'C:\GHUFRAN\Old\PythonScripting\Niagara\Evap_Cooler_Points_1.xlsx'
)