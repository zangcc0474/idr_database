from auxiliary import *

CF_Name = "CF 1"
Quarter_Name = "Q32018"
file_location = "C:/Users/re70296/Downloads/IDR Database/Database and Scripts/IDR_file.xls"
# Open the workbook and define the worksheet
book = xlrd.open_workbook(file_location)

SHEETS_index = {'IDR Quarterly Inputs': 'IDRQuarterlyInputs','IDR Input - NFI-ODCE Index': 'IDRIpNFIODCE',
'IDR Input - NFI-ODCE X Index': 'IDRIpNFIODCEX','Input - CF Fund Level (NAV)': 'FundLevelNAV',
'Input-CF Diversification (NAV)': 'DiversificationNAV','Input - CF Portfolio':'Portfolio',
'Input - CF Acquisitions': 'Acquisitions','Input - CF Dispositions': 'Dispositions',
'Input - CF Performance': 'Performance','Input - CF Disclosures': 'Disclosures'}

database, cursor = create_connection()

cfid_curr, quarterid_curr = create_strong_tables(CF_Name = CF_Name, Quarter_Name = Quarter_Name, database = database, cursor = cursor, )

###Update all tables. Insert Script here
for sheetname in SHEETS_index:
	insert_values(book, cfid_curr, quarterid_curr, sheetname, SHEETS_index, database, cursor, file_location)

close_database(database, cursor)