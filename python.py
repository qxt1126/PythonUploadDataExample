import xlrd
import MySQLdb

xlsFile = xlrd.open_workbook("data.xls")
dataSheet = xlsFile.sheet_by_name("data1")

database = MySQLdb.connect (host="localhost", user = "root", passwd = "", db = "mysqlPython")

cursor = database.cursor()

query = """INSERT INTO Cases (CaseID, CaseStatusEffectiveDate, ServiceRequisitionID, ServiceRequisitionAmount, ServiceProviderName, AlternateName, TargetGroup) VALUES (%s, %s, %s, %s, %s, %s, %s)"""

for r in range(1, dataSheet.nrows):
		CaseID = dataSheet.cell(r,).value
		CaseStatusEffectiveDate	= dataSheet.cell(r,1).value
		ServiceRequisitionID = dataSheet.cell(r,2).value
		ServiceRequisitionAmount = dataSheet.cell(r,3).value
		ServiceProviderName	= dataSheet.cell(r,4).value
		AlternateName = dataSheet.cell(r,5).value
		TargetGroup	= dataSheet.cell(r,6).value

		values = (CaseID, CaseStatusEffectiveDate, ServiceRequisitionID, ServiceRequisitionAmount, ServiceProviderName, AlternateName, TargetGroup)

		cursor.execute(query, values)

cursor.close()

database.commit()

database.close()

columns = str(dataSheet.ncols)
rows = str(dataSheet.nrows)
print("Finished!")