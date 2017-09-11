'..............................................................................................................................

'Test Name : This test will set all the preconditions for the test case.

'..............................................................................................................................


'Regression testing connection:
Set WshShell = CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("User")
	Set EPS2DBObject = CreateObject("ADODB.Connection") 
	Dim    vSchema 	, vEnvironment, vDSN
	vSchema  =  WshSysEnv("epsvSchema")
	vEnvironment = WshSysEnv("epsEnvironment")
	vDSN               =  WshSysEnv("vDSN")	
	dbpassword = WshSysEnv("vDbpwd")
	dbuser = WshSysEnv("vDbuser")
	serverip = WshSysEnv("vServerip")
	releasebase = WshSysEnv("vRELEASE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

'Import values from sheet

ndc1 = DataTable.Value("NDC","Global")

'Set single drug batch processing to Yes

SDBPEnabled("Yes")

Wait(Min_Wait)

'Set all values in single drug batch processing on application settings

SDBPApplicationSettings_2609 "No","No","No","No","No","No","No","No","No","Yes","Yes"

Wait(Min_Wait)

'Check SDBP on drug

SDBPDrug("ON")

Wait(Min_Wait)

'Check Immnunization checkbox on Drug>Additional Screen

DrugShareImmunization ndc1,"ON"

Reporter.ReportEvent micDone,"Preconditions","All preconditions have been set"
