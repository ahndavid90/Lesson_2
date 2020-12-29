import pandas as pd
import openpyxl
import os

def automation():
    print('start!')
    print('test')
    print('test2')
    original_path = os.path.join(os.getcwd(),'Summary_12-23-2020.xlsx')
    rates_path = os.path.join(os.getcwd(), 'Inputs_12-24-2020.xlsx')

    #Creates a df - treasury rates
    df_rates = pd.read_excel(rates_path, sheet_name='Rates', engine='openpyxl', index_col=0)

    #Sets val_date
    #This is due to the way datetime values are stored in pandas: using the numpy datetime64[ns] dtype. 
    val_date = df_rates.columns[0].strftime('%m-%d-%Y')

    print(df_rates)
    print(val_date)
    #Opens the summary_sheet
    summary_workbook = openpyxl.load_workbook(filename = original_path)
    input_sheet = summary_workbook['Inputs']

    input_sheet['B2'].value = val_date

    row_max = len(df_rates.index)

    for i in range(0, row_max):
        #print(i)
        #print(df_rates.iloc[i, 0])          #row, column
        input_sheet.cell(i + 9, 3).value = df_rates.iloc[i, 0]      #row, columns

    destination = os.path.join(os.getcwd(), r'Summary_{}.xlsx'.format(val_date))

    summary_workbook.save(filename = destination)
    print('workbook saved as {}'.format(destination))
    print('Success!')

#Set special variable before running the code
#If this python file is being imported
if __name__ == '__main__':
    automation()

