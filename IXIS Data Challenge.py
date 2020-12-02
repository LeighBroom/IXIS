import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter

#Read in the Session Count CSV and preform intitial data exploration
df = pd.read_csv('''C:/Users/broom/Downloads/DataAnalyst_Ecom_data_sessionCounts.csv''')
df2 = pd.read_csv('''C:/Users/broom/Downloads/DataAnalyst_Ecom_data_addsToCart.csv''')
#Convert date strings to datetime
df['dim_date'] = pd.to_datetime(df['dim_date'])
#Check to see the composition of the first few rows
print(df.head())
#Check for null values
print(df.isnull().any())


# print(df.groupby(['dim_browser']).sum().sort_values('sessions', ascending=False))
# By using group by we get essentially a pivot table. A quick summary of the data
df_groupby = df.groupby(['dim_browser']).sum().sort_values('sessions', ascending=False)
df_groupby['t/s'] = df_groupby['transactions']/df_groupby['sessions']
df_groupby['q/s'] = df_groupby['QTY']/df_groupby['sessions']
df_groupby['q/t'] = df_groupby['QTY']/df_groupby['transactions']

#Turn anything divided by 0 back to 0 instead of infinity or not a number.
df_groupby = df_groupby.replace([np.inf, -np.inf, np.NaN], 0)
df_groupby[['t/s', 'q/s', 'q/t']]

#Preparing sheet one. Data
df_sheet1_columns = df[['dim_date', 'dim_deviceCategory', 'sessions', 'transactions', 'QTY']]

df_sheet1_columns = df_sheet1_columns.replace([np.inf, -np.inf, np.NaN], 0)
df_month_x_device = df_sheet1_columns.groupby([pd.Grouper(key='dim_date', freq='M'), pd.Grouper('dim_deviceCategory')]).sum()
#Reset Index can be left out if the visual of this is needed to be cleaner as opposed to used for processing
df_month_x_device = df_month_x_device.reset_index().sort_values(['dim_date', 'dim_deviceCategory'], ascending=False).reset_index(drop=True)
#Organizing values for chart
df_month_x_device = df_month_x_device.sort_values(['dim_deviceCategory', 'dim_date']).reset_index(drop=True)
#Calculating ECR after aggregation so the results aren't summed
df_month_x_device['ECR'] = df_month_x_device['transactions']/df_month_x_device['sessions']
df_month_x_device['dim_month'] = df_month_x_device['dim_date'].dt.strftime('%B %Y')


writer = pd.ExcelWriter('C:/Users/broom/Documents/IXIS Data Challenge.xlsx', engine='xlsxwriter')
workbook = writer.book
df_month_x_device.to_excel(writer, sheet_name='Sheet1')

#Excel formatting
worksheet1 = writer.sheets['Sheet1']
worksheet1.set_column('A:A', 2.5)
worksheet1.set_column('B:B', 18)
worksheet1.set_column('C:C', 19)
worksheet1.set_column('D:D', 8)
worksheet1.set_column('E:E', 12)
worksheet1.set_column('F:F', 8)
worksheet1.set_column('G:G', 9)
worksheet1.set_column('H:H', 15)

#Sessions Chart
chart1 = workbook.add_chart({'type': 'line'})
#Seires lists explained [Name of sheet, first row, first column, last row, last column]
chart1.add_series({
    'name': 'Desktop',
    'categories': ['Sheet1', 1, 7, 12, 7],
    'values': ['Sheet1', 1, 3, 12, 3]})
chart1.add_series({
    'name': 'Mobile',
    'categories': ['Sheet1', 13, 7, 24, 7],
    'values': ['Sheet1', 13, 3, 24, 3]})
chart1.add_series({
    'name': 'Tablet',
    'categories': ['Sheet1', 25, 7, 36, 7],
    'values': ['Sheet1', 25, 3, 36, 3]})
chart1.set_title({'name': 'Sessions By Device'})
worksheet1.insert_chart('I2', chart1)

#Transaction Chart
#just change values to transactions, so column 3 -> 4
chart2 = workbook.add_chart({'type': 'line'})
chart2.add_series({
    'name': 'Desktop',
    'categories': ['Sheet1', 1, 7, 12, 7],
    'values': ['Sheet1', 1, 4, 12, 4]})
chart2.add_series({
    'name': 'Mobile',
    'categories': ['Sheet1', 13, 7, 24, 7],
    'values': ['Sheet1', 13, 4, 24, 4]})
chart2.add_series({
    'name': 'Tablet',
    'categories': ['Sheet1', 25, 7, 36, 7],
    'values': ['Sheet1', 25, 4, 36, 4]})
chart2.set_title({'name': 'Transactions By Device'})
worksheet1.insert_chart('I17', chart2)

#QTY Chart
#For efficiency of space, all charts can be done in a for loop if needed
chart3 = workbook.add_chart({'type': 'line'})
chart3.add_series({
    'name': 'Desktop',
    'categories': ['Sheet1', 1, 7, 12, 7],
    'values': ['Sheet1', 1, 5, 12, 5]})
chart3.add_series({
    'name': 'Mobile',
    'categories': ['Sheet1', 13, 7, 24, 7],
    'values': ['Sheet1', 13, 5, 24, 5]})
chart3.add_series({
    'name': 'Tablet',
    'categories': ['Sheet1', 25, 7, 36, 7],
    'values': ['Sheet1', 25, 5, 36, 5]})
chart3.set_title({'name': 'QTY By Device'})
worksheet1.insert_chart('Q2', chart3)

#ECR Chart
chart4 = workbook.add_chart({'type': 'line'})
chart4.add_series({
    'name': 'Desktop',
    'categories': ['Sheet1', 1, 7, 12, 7],
    'values': ['Sheet1', 1, 6, 12, 6]})
chart4.add_series({
    'name': 'Mobile',
    'categories': ['Sheet1', 13, 7, 24, 7],
    'values': ['Sheet1', 13, 6, 24, 6]})
chart4.add_series({
    'name': 'Tablet',
    'categories': ['Sheet1', 25, 7, 36, 7],
    'values': ['Sheet1', 25, 6, 36, 6]})
chart4.set_title({'name': 'ECR By Device'})
worksheet1.insert_chart('Q17', chart4)

#Preparing sheet two
df_sheet2 = df[['dim_date', 'dim_deviceCategory', 'sessions', 'transactions', 'QTY']][df['dim_date'] > '2013-04-30']
df_sheet2['ECR'] = df_sheet2['transactions']/df_sheet2['sessions']
df_sheet2 = df_sheet2.replace([np.inf, -np.inf, np.NaN], 0)
#Instead of reseting index like before we're transposing the dataframe to use the months as columns. 
#Since the months are already indexes after a groupby, just transpose immediately 
df_month_over_month = df_sheet2.groupby(pd.Grouper(key='dim_date',freq='M')).sum().T
#Fixing ECR row after aggregation so it's correctly represented
#Columns[0] is may and columns[1] is june
#so ecr for may is reset as transactions for may over sessions for may etc.
df_month_over_month.at[df_month_over_month.index[3], df_month_over_month.columns[0]] = df_month_over_month.iloc[1][0]/df_month_over_month.iloc[0][0]
df_month_over_month.at[df_month_over_month.index[3], df_month_over_month.columns[1]] = df_month_over_month.iloc[1][1]/df_month_over_month.iloc[0][1]


#By appending with .loc we don't change the index
df_month_over_month.loc['addsToCart'] = [df2[(df2['dim_year'] == 2013) & (df2['dim_month'] == 5)]['addsToCart'].values[0],
                                        df2[(df2['dim_year'] == 2013) & (df2['dim_month'] == 6)]['addsToCart'].values[0]]

#Doing a calculation will change the column types from datetime to object, so make sure to use correct ones after
df_month_over_month['Total Change'] = df_month_over_month['2013-06-30']-df_month_over_month['2013-05-31']
df_month_over_month['Percent Change'] = df_month_over_month[df_month_over_month.columns[0]]/(df_month_over_month[df_month_over_month.columns[1]] - df_month_over_month[df_month_over_month.columns[0]])
#This for loop handles any fractions. If the total change is less than 1 and more than -1 multiplication should be used for the percent, not division.
for index, i in df_month_over_month['Total Change'].iteritems():
    if i > -1 and i < 1:
        df_month_over_month.at[index, 'Percent Change'] = (df_month_over_month.at[index, df_month_over_month.columns[0]] * i)

df_month_over_month.to_excel(writer, sheet_name='Sheet2')
worksheet2 = writer.sheets['Sheet2']
worksheet2.set_column('A:A', 11.2)
worksheet2.set_column('B:B', 17.71)
worksheet2.set_column('C:C', 17.71)
worksheet2.set_column('D:D', 12.41)
worksheet2.set_column('E:E', 14.21)

chart5 = workbook.add_chart({'type': 'column'})
chart5.add_series({'name': 'May', 'categories': 'Sessions', 'values': ['Sheet2', 1, 1, 1, 1]})
chart5.add_series({'name': 'June', 'categories': 'Sessions', 'values': ['Sheet2', 1, 2, 1, 2]})
chart5.set_title({'name': 'Sessions Per Month'})
chart5.set_x_axis({'display_units_visible': False})
worksheet2.insert_chart('G2', chart5)

chart6 = workbook.add_chart({'type': 'column'})
chart6.add_series({'name': 'May', 'categories': 'Sessions', 'values': ['Sheet2', 2, 1, 2, 1]})
chart6.add_series({'name': 'June', 'categories': 'Sessions', 'values': ['Sheet2', 2, 2, 2, 2]})
chart6.set_title({'name': 'Transactions Per Month'})
worksheet2.insert_chart('G17', chart6)

chart7 = workbook.add_chart({'type': 'column'})
chart7.add_series({'name': 'May', 'categories': 'Sessions', 'values': ['Sheet2', 3, 1, 3, 1]})
chart7.add_series({'name': 'June', 'categories': 'Sessions', 'values': ['Sheet2', 3, 2, 3, 2]})
chart7.set_title({'name': 'QTY Per Month'})
worksheet2.insert_chart('O2', chart7)

worksheet2.write('A9', 'Adds To Cart')
addsData = [i for i in df2['addsToCart'].values]
worksheet2.write_column('A10', addsData)
addsMonths = ['July 12', 'August 12', 'September 12', 'October 12', 'November 12', 'December 12',
              'January 13', 'February 13', 'March 13', 'April 13', 'May 13', 'June 13']
worksheet2.write_column('B9', ['Months'])
worksheet2.write_column('B10', addsMonths)

chart8 = workbook.add_chart({'type': 'line'})
chart8.add_series({'name': 'Adds To Cart', 'categories': ['Sheet2', 9, 1, 20, 1], 'values': ['Sheet2', 9, 0, 20, 0]})
worksheet2.insert_chart('O17', chart8)

writer.save()