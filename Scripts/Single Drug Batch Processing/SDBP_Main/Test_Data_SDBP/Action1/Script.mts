'......................................................................................................................................

'Test Name : Test_Data_SDBP

'Test Description : This test will setup the data for the test case to run.

'Date Modified : 7 August 2017

'......................................................................................................................................
 @@ hightlight id_;_20590695_;_script infofile_;_ZIP::ssf181.xml_;_
 
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
	WF_mode = WshSysEnv("vMODE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

'Import variables from sheet

thirdparty1 = DataTable.Value("ThirdParty1","Global")
thirdparty2 = DataTable.Value("ThirdParty2","Global")
thirdparty3 = DataTable.Value("ThirdParty3","Global")
cardholderid1 = DataTable.Value("CARDHOLDERID1","Global")
cardholderid2 = DataTable.Value("CARDHOLDERID2","Global")
cardholderid3 = DataTable.Value("CARDHOLDERID3","Global")

'Call actions here

RunAction "Action1 [Create_Patient]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Create_Prescriber]", oneIteration
Wait(Iteration_Wait)
AddNewTP_Prescriber(thirdparty1)
Wait(Iteration_Wait)
AddNewTP_Prescriber(thirdparty2)
Wait(Iteration_Wait)
AddNewTP_Prescriber(thirdparty3)
Wait(Iteration_Wait)
CardholderFormat_TP(thirdparty1)
Wait(Iteration_Wait)
CardholderFormat_TP(thirdparty2)
Wait(Iteration_Wait)
CardholderFormat_TP(thirdparty3)
Wait(Iteration_Wait)
AddNewTP_Patient thirdparty1,cardholderid1
Wait(Iteration_Wait)
AddNewTP_Patient thirdparty2,cardholderid2
Wait(Iteration_Wait)
AddNewTP_Patient thirdparty3,cardholderid3

Reporter.ReportEvent micDone,"Test Data","Test Data for this test has been added successfully"
