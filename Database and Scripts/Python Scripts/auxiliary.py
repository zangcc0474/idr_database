import xlrd
import pymysql as MySQLdb

import numpy as np
import pandas as pd

def create_connection():
	# Establish a MySQL connection
	database = MySQLdb.connect(host="localhost", user = "root", passwd = "Tanmay9830!", db = "idr_db")

	# Get the cursor, which is used to traverse the database, line by line
	cursor = database.cursor()
	cursor.execute("USE idr_db") # select the database
	cursor.execute("SET SESSION sql_mode = ''")
	database.commit()
	print('Database Connection Opened')
	return database, cursor

def return_pk_index(tablename, database, cursor, pkid='PKID'):
	'''To get the maximum primary key for a specific table'''
	query_pk = """SELECT MAX(PKID) FROM """ + tablename
	cursor.execute(query_pk)
	query_pk_result = cursor.fetchone()[0]

	if(query_pk_result):
		return query_pk_result
	else:
		return 0
	#return 0

def create_strong_tables(CF_Name, Quarter_Name, database, cursor):
	'''
	Checking if corresponding Quarter and Component Fund are already present in the database
	'''
	cursor.execute("SELECT Name FROM Quarter WHERE name=%s", Quarter_Name)
	exists = cursor.fetchone()
	if(exists):
		cursor.execute("SELECT QuarterID FROM Quarter WHERE name=%s", Quarter_Name)
		quarterid_curr = cursor.fetchone()[0]
	else:
		cursor.execute("SELECT MAX(QuarterID) + 1 FROM Quarter")
		quarterid_curr = cursor.fetchone()[0]
		if(quarterid_curr):
			print("QuarterID: " + str(quarterid_curr))
		else:
			print("No entries in the Quarter table. Creating one now.")
			quarterid_curr = 1
		quarter_query = """INSERT INTO Quarter (quarterid, name) VALUES (%s, %s)"""
		quarter_values = (quarterid_curr, Quarter_Name)
		cursor.execute(quarter_query, quarter_values)

	cursor.execute("SELECT Name FROM ComponentFund WHERE name=%s", CF_Name)
	exists = cursor.fetchone()
	if(exists):
		cursor.execute("SELECT CFID FROM ComponentFund WHERE name=%s", CF_Name)
		cfid_curr = cursor.fetchone()[0]
	else:
		cursor.execute("SELECT MAX(CFID) + 1 FROM ComponentFund")
		cfid_curr = cursor.fetchone()[0]
		if(cfid_curr):
			print("CFID: " + str(cfid_curr))
		else:
			print("No entries in the ComponentFund table. Creating one now.")
			cfid_curr = 1
		cfid_query = """INSERT INTO ComponentFund (CFID, name) VALUES (%s, %s)"""
		cfid_values = (cfid_curr, CF_Name)
		cursor.execute(cfid_query, cfid_values)

	return cfid_curr, quarterid_curr

def insert_values(book, cfid_curr, quarterid_curr, sheetname, SHEETS_index, database, cursor, file_location):
	'''
	Code to insert values
	'''
	if(SHEETS_index.get(sheetname)=='FundLevelNAV'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, FieldName, Value)
 		VALUES (%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)

		for r in range(1, sheet.nrows):
			if(r!=27):
				PKID = PKID + 1
				CFID = cfid_curr
				QuarterID = quarterid_curr

				if(r<27):
					FieldName = sheet.cell(1,0).value + ": " + sheet.cell(r,0).value
				elif(r>27):
					FieldName = sheet.cell(27,0).value + ": " + sheet.cell(r,0).value

				Value = sheet.cell(r,1).value

				# Assign values from each row
				values = (PKID, CFID, QuarterID, FieldName, Value)

				# Execute sql Query
				cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='Acquisitions'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, InvestmentName, Streetaddress, msa, city, state, zipcode, size, propertytype, 
		acquisitiondate, propertylifecycle, occupancy, totalacquicostamt, totalequityamt, totaldebtamt, entrycapratepc, total1yrgoingfwdnoi, 
		fundownershippc, proformairr, proformamultiple)
 		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)
		df = pd.read_excel(file_location, sheetname, header = None)
		last_row = int(df.isnull().all(1).nonzero()[0][0])

		for r in range(1, last_row):
			if (r!=27):
				PKID = PKID + 1
				CFID = cfid_curr
				QuarterID = quarterid_curr
				InvestmentName = sheet.cell(r,0).value
				Streetaddress = sheet.cell(r,1).value
				msa = sheet.cell(r,2).value
				city = sheet.cell(r,3).value
				state = sheet.cell(r,4).value
				zipcode = sheet.cell(r,5).value
				size = sheet.cell(r,6).value
				propertytype = sheet.cell(r,7).value
				acquisitiondate = sheet.cell(r,8).value
				propertylifecycle = sheet.cell(r,9).value
				occupancy = sheet.cell(r,10).value
				totalacquicostamt = sheet.cell(r,11).value
				totalequityamt = sheet.cell(r,12).value
				totaldebtamt = sheet.cell(r,13).value
				entrycapratepc = sheet.cell(r,14).value
				total1yrgoingfwdnoi = sheet.cell(r,15).value
				fundownershippc = sheet.cell(r,16).value
				proformairr = sheet.cell(r,17).value
				proformamultiple = sheet.cell(r,18).value
				# Assign values from each row
				values = (PKID, CFID, QuarterID, InvestmentName, Streetaddress, msa, city, state, zipcode, size, propertytype, 
						acquisitiondate, propertylifecycle, occupancy, totalacquicostamt, totalequityamt, totaldebtamt, entrycapratepc, total1yrgoingfwdnoi, 
						fundownershippc, proformairr, proformamultiple)

				# Execute sql Query
				cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='Disclosures'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, Keypersonnelhired, Keypersonneldeparted, investmentguidelines, 
		nfiodceguidelines, litigation, otherevents, descriptiondispositions, descriptionacquisitions)
 		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)

		PKID = PKID + 1
		CFID = cfid_curr
		QuarterID = quarterid_curr
		Keypersonnelhired = sheet.cell(0,1).value
		Keypersonneldeparted = sheet.cell(1,1).value
		investmentguidelines = sheet.cell(2,1).value
		nfiodceguidelines = sheet.cell(3,1).value
		litigation = sheet.cell(4,1).value
		otherevents = sheet.cell(5,1).value
		descriptiondispositions = sheet.cell(6,1).value
		descriptionacquisitions = sheet.cell(7,1).value

		values = (PKID, CFID, QuarterID, Keypersonnelhired, Keypersonneldeparted, investmentguidelines, nfiodceguidelines, 
				litigation, otherevents, descriptiondispositions, descriptionacquisitions)

		# Execute sql Query
		cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='Dispositions'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)
		print(sheet.nrows)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, InvestmentName, Streetaddress, msa, city, state, zipcode, size, propertytype, acquisitiondate, 
		dispositiondate, holdingperiod, propertylifecycle, occupancy, totalcostbasisamt, totalsalepriceamt, previousquartergav, 
		totalrealizedproceedsamt, netrealizedproceedsamt, exitcapratepc, realizedgrossirrpc, realizedmultiple)
 		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)
		df = pd.read_excel(file_location, sheetname, header = None)
		last_row = int(df.isnull().all(1).nonzero()[0][0])

		for r in range(1, last_row):
			PKID = PKID + 1
			CFID = cfid_curr
			QuarterID = quarterid_curr
			InvestmentName = sheet.cell(r,0).value
			Streetaddress = sheet.cell(r,1).value
			msa = sheet.cell(r,2).value
			city = sheet.cell(r,3).value
			state = sheet.cell(r,4).value
			zipcode = sheet.cell(r,5).value
			size = sheet.cell(r,6).value
			propertytype = sheet.cell(r,7).value
			acquisitiondate = sheet.cell(r,8).value
			dispositiondate = sheet.cell(r,9).value
			holdingperiod  = sheet.cell(r,10).value
			propertylifecycle = sheet.cell(r,11).value
			occupancy = sheet.cell(r,12).value
			totalcostbasisamt = sheet.cell(r,13).value
			totalsalepriceamt = sheet.cell(r,14).value
			previousquartergav = sheet.cell(r,15).value
			totalrealizedproceedsamt = sheet.cell(r,16).value
			netrealizedproceedsamt = sheet.cell(r,17).value
			exitcapratepc = sheet.cell(r,18).value
			realizedgrossirrpc = sheet.cell(r,19).value
			realizedmultiple = sheet.cell(r,20).value
			# Assign values from each row
			values = (PKID, CFID, QuarterID, InvestmentName, Streetaddress, msa, city, state, zipcode, size, propertytype, acquisitiondate, dispositiondate, 
				holdingperiod, propertylifecycle, occupancy, totalcostbasisamt, totalsalepriceamt, previousquartergav, totalrealizedproceedsamt, 
				netrealizedproceedsamt, exitcapratepc, realizedgrossirrpc, realizedmultiple)

			# Execute sql Query
			cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='DiversificationNAV'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, FieldName, Totalpc, Totalamt, office, multifamily, industrial, retail, hotel, 
		healthcare, storage, studenthousing, seniorliving, parking, land, other)
 		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)

		for r in range(2, sheet.nrows):
			if((r!=11) & (r!=21) & (r!=25)):
				PKID = PKID + 1
				CFID = cfid_curr
				QuarterID = quarterid_curr

				if(r<11):
					FieldName = sheet.cell(1,0).value + ": " + sheet.cell(r,0).value
				elif((r>11) & (r<21)):
					FieldName = sheet.cell(11,0).value + ": " + sheet.cell(r,0).value
				elif((r>21) & (r<25)):
					FieldName = sheet.cell(21,0).value + ": " + sheet.cell(r,0).value
				elif(r>25):
					FieldName = sheet.cell(21,0).value + ": " + sheet.cell(r,0).value

				Totalpc = sheet.cell(r,1).value
				Totalamt = sheet.cell(r,2).value
				office = sheet.cell(r,3).value
				multifamily = sheet.cell(r,4).value
				industrial = sheet.cell(r,5).value
				retail = sheet.cell(r,6).value
				hotel = sheet.cell(r,7).value
				healthcare = sheet.cell(r,8).value
				storage = sheet.cell(r,9).value
				studenthousing = sheet.cell(r,10).value
				seniorliving = sheet.cell(r,11).value
				parking = sheet.cell(r,12).value
				land = sheet.cell(r,13).value
				other = sheet.cell(r,14).value

				# Assign values from each row
				values = (PKID, CFID, QuarterID, FieldName, Totalpc, Totalamt, office, multifamily, industrial, retail, hotel, healthcare, storage, 
					studenthousing, seniorliving, parking, land, other)

				# Execute sql Query
				cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='IDRIpNFIODCE'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, FieldName, ValueType1, ValueType2)
 		VALUES (%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)

		for r in range(2, 86):
			if((r!=23) & (r!=33) & (r!=37) & (r!=47)& (r!=73)):
				PKID = PKID + 1
				CFID = cfid_curr
				QuarterID = quarterid_curr

				if(r<23):
					FieldName = sheet.cell(1,0).value + ": " + sheet.cell(r,0).value
				elif((r>23) & (r<33)):
					FieldName = sheet.cell(23,0).value + ": " + sheet.cell(r,0).value
				elif((r>33) & (r<37)):
					FieldName = sheet.cell(33,0).value + ": " + sheet.cell(r,0).value
				elif((r>37) & (r<47)):
					FieldName = sheet.cell(37,0).value + ": " + sheet.cell(r,0).value
				elif((r>47) & (r<73)):
					FieldName = sheet.cell(47,0).value + ": " + sheet.cell(r,0).value
				elif(r>73):
					FieldName = sheet.cell(73,0).value + ": " + sheet.cell(r,0).value

				ValueType1 = sheet.cell(r,1).value
				ValueType2 = sheet.cell(r,2).value
				# Assign values from each row
				values = (PKID, CFID, QuarterID, FieldName, ValueType1, ValueType2)

				# Execute sql Query
				cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='IDRIpNFIODCEX'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, FieldName, ValueType1, ValueType2)
 		VALUES (%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)

		for r in range(2, 86):
			if((r!=23) & (r!=33) & (r!=37) & (r!=47)& (r!=73)):
				PKID = PKID + 1
				CFID = cfid_curr
				QuarterID = quarterid_curr

				if(r<23):
					FieldName = sheet.cell(1,0).value + ": " + sheet.cell(r,0).value
				elif((r>23) & (r<33)):
					FieldName = sheet.cell(23,0).value + ": " + sheet.cell(r,0).value
				elif((r>33) & (r<37)):
					FieldName = sheet.cell(33,0).value + ": " + sheet.cell(r,0).value
				elif((r>37) & (r<47)):
					FieldName = sheet.cell(37,0).value + ": " + sheet.cell(r,0).value
				elif((r>47) & (r<73)):
					FieldName = sheet.cell(47,0).value + ": " + sheet.cell(r,0).value
				elif(r>73):
					FieldName = sheet.cell(73,0).value + ": " + sheet.cell(r,0).value

				ValueType1 = sheet.cell(r,1).value
				ValueType2 = sheet.cell(r,2).value
				# Assign values from each row
				values = (PKID, CFID, QuarterID, FieldName, ValueType1, ValueType2)

				# Execute sql Query
				cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='Performance'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, Timeweightedreturns, qtr, 1year, 3year, 5year, 7year, 10year, sinceinception)
 		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)

		for r in range(1, sheet.nrows):
			PKID = PKID + 1
			CFID = cfid_curr
			QuarterID = quarterid_curr
			Timeweightedreturns = sheet.cell(r,0).value
			qtr = sheet.cell(r,1).value
			year_1 = sheet.cell(r,2).value
			year_3 = sheet.cell(r,3).value
			year_5 = sheet.cell(r,4).value
			year_7 = sheet.cell(r,5).value
			year_10 = sheet.cell(r,6).value
			sinceinception = sheet.cell(r,7).value
			# Assign values from each row
			values = (PKID, CFID, QuarterID, Timeweightedreturns, qtr, year_1, year_3, year_5, year_7, year_10, sinceinception)
			# Execute sql Query
			cursor.execute(query, values)

		# Commit the transaction
		database.commit()

	elif(SHEETS_index.get(sheetname)=='Portfolio'):
		tablename = SHEETS_index.get(sheetname)
		sheet = book.sheet_by_name(sheetname)
		delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr)

		# Create the INSERT INTO sql query
		query = """INSERT INTO """ + tablename + """ (PKID, CFID, QuarterID, InvestmentName, Streetaddress, msa, city, state, zipcode, size, propertytype, 
		acquisitiondate, propertylifecycle, occupancy, totalcostbasisamt, totalgavamt, totaldebtamt, totalnav, total1yrgoingfwdnoi, currentcapratepc, 
		ownershippc, fundgav, funddebt, fundnav, stabilizedvalue)
 		VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""

		PKID = return_pk_index(tablename, database, cursor)
		df = pd.read_excel(file_location, sheetname, header = None)
		last_row = int(df.isnull().all(1).nonzero()[0][0])

		for r in range(1, last_row):
			PKID = PKID + 1
			CFID = cfid_curr
			QuarterID = quarterid_curr
			InvestmentName = sheet.cell(r,0).value
			Streetaddress = sheet.cell(r,1).value
			msa = sheet.cell(r,2).value
			city = sheet.cell(r,3).value
			state = sheet.cell(r,4).value
			zipcode = sheet.cell(r,5).value
			size = sheet.cell(r,6).value
			propertytype = sheet.cell(r,7).value
			acquisitiondate = sheet.cell(r,8).value
			propertylifecycle = sheet.cell(r,9).value
			occupancy = sheet.cell(r,10).value
			totalcostbasisamt = sheet.cell(r,11).value
			totalgavamt = sheet.cell(r,12).value
			totaldebtamt = sheet.cell(r,13).value
			totalnav = sheet.cell(r,14).value
			total1yrgoingfwdnoi = sheet.cell(r,15).value
			currentcapratepc = sheet.cell(r,16).value
			ownershippc = sheet.cell(r,17).value
			fundgav = sheet.cell(r,18).value
			funddebt = sheet.cell(r,19).value
			fundnav = sheet.cell(r,20).value
			stabilizedvalue = sheet.cell(r,21).value
			# Assign values from each row
			values = (PKID, CFID, QuarterID, InvestmentName, Streetaddress, msa, city, state, zipcode, size, propertytype, acquisitiondate, 
			propertylifecycle, occupancy, totalcostbasisamt, totalgavamt, totaldebtamt, totalnav, total1yrgoingfwdnoi, currentcapratepc, 
			ownershippc, fundgav, funddebt, fundnav, stabilizedvalue)

			# Execute sql Query
			cursor.execute(query, values)

		# Commit the transaction
		database.commit()

def delete_duplicates(database, cursor, tablename, cfid_curr, quarterid_curr):
	# cursor.execute("""DELETE FROM %s WHERE CFID=%s;""", (tablename, cfid_curr))
	# database.commit()
	# cursor.execute("""DELETE FROM %s WHERE QuarterID=%s;""", (tablename, quarterid_curr))
	# database.commit()
	print('In table '+ tablename)

def close_database(database, cursor):
	# Close the cursor
	cursor.close()

	# Close the database connection
	database.close()
	print('Database Connection Closed')