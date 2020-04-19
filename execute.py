# import libraries
from pandas import read_excel, merge, ExcelWriter
import xlsxwriter

# define file names, files must be named as follows
file1 = "merge-file-1.xlsx"
file2 = "merge-file-2.xlsx"

# get pandas dataframes from files
df1 = read_excel(file1)
df2 = read_excel(file2)

# capitalize both column "Device Category" from df1 and column "DeviceType" from df2
# so that they match regardless of the capitalization
df1["Device Category"] = df1["Device Category"].str.lower() 
df1["Device Category"] = df1["Device Category"].str.capitalize() 
df2["DeviceType"] = df2["DeviceType"].str.lower()
df2["DeviceType"] = df2["DeviceType"].str.capitalize()

# replace the "Computer" keyword with "Dsktop" on "DeviceType" column of df2
df2["DeviceType"] = df2["DeviceType"].replace("Computer", "Desktop")

# inner merge 
df = merge(df1, df2, left_on = ["Page", "Device Category"], right_on = ["Labels", "DeviceType"], how='inner')

# take the columns that are needed
# Page, Campaign, Ad group, Device Category, Publisher CTR, Avg. CPC, Clicks, Publisher Clicks, Spend, Publisher Revenue, Profit, Margin
df = df[['Page', 'Campaign', 'Ad group', 'Device Category', 'Publisher CTR', 'Avg. CPC', 'Clicks', 'Publisher Clicks', 'Spend', 'Publisher Revenue']]

# add two more columns based on a calculation from other columns
df['Profit'] = df['Publisher Revenue'] - df['Spend']
df['Margin'] = df['Profit'] / df['Publisher Revenue']
df['Margin'] = df['Margin'][df['Margin'] <= 100]
df['Margin'] = df['Margin'][df['Margin'] >= 0]
        
# replace -inf/nan values (string) with $0.00 
df['Margin'] = df['Margin'].astype(str).str.replace('-inf','0')
df['Margin'] = df['Margin'].astype(str).str.replace('nan','0')
df['Margin'] = df['Margin'].astype(float)

# rename the column names as follows:
# from this - current
# Page	Campaign	Ad group	Device Category	Publisher CTR	Avg. CPC	Clicks	Publisher Clicks	Spend	Publisher Revenue	Profit	Margin
# to this - requested
# Page	Campaign	Ad group	Device	Adsense CTR	Bing CPC	Bing Clicks	Adsense Clicks	Bing Spend	Adsense Revenue	Profit	Margin
df.columns = ['Page', 'Campaign', 'Ad group', 'Device', 'Adsense CTR', 'Bing CPC', 'Bing Clicks', 'Adsense Clicks', 'Bing Spend', 'Adsense Revenue', "Profit", "Margin"]          
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = ExcelWriter("result.xlsx", engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, index=False, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add some cell formats.
number_format = workbook.add_format({'num_format': '#,##0'})
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
percentage_format = workbook.add_format({'num_format': '0.00%'})
column_width = 18

# Set the column width and format.
# number formatL
worksheet.set_column('A:A', column_width*2, None)
worksheet.set_column('B:B', column_width, None)
worksheet.set_column('C:C', column_width*2, None)
worksheet.set_column('D:D', column_width, None)
worksheet.set_column('E:E', column_width, percentage_format)
worksheet.set_column('F:F', column_width, currency_format)
worksheet.set_column('G:G', column_width, number_format)
worksheet.set_column('H:H', column_width, number_format)
worksheet.set_column('I:I', column_width, currency_format)
worksheet.set_column('J:J', column_width, currency_format)
worksheet.set_column('K:K', column_width, currency_format)
worksheet.set_column('L:L', column_width, percentage_format)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
