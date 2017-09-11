'............................................................................................................................

'Test Name : Verify Password about to expire PopUp to Display on EPS home page when Authentication Mode is 1,Mode is 2 and when client inactive session time out is '60'

'Test Description : Verify Password about to expire PopUp to Display on EPS home page when Authentication Mode is 1,Mode is 2 and when client inactive session time out is '60'

'TestRail ID : C200227

'Author : Kashish Ambwani

'Date Modified : 1 August 2017

'............................................................................................................................

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
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

WshSysEnv ("WF") = "C200227"

'Import data sheet

importDataSheet()

'Call actions here

RunAction "Action1 [C200227_AT_RF]", oneIteration

Reporter.ReportEvent micDone,"C200227","Test case execution has been completed"
