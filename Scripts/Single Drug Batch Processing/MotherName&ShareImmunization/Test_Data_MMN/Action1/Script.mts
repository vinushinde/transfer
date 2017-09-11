'......................................................................................................................................

'Test Name : Test_Data_SDBP

'Test Description : This test will setup the data for the test case to run.

'Date Modified : 9 August 2017

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

'Call actions here

RunAction "Action1 [Create_Patient]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Create_Patient2]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Create_Patient3]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Create_Prescriber]", oneIteration


Reporter.ReportEvent micDone,"Test Data","Test Data for this test has been added successfully"
