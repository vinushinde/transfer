'................................................................................................................................

'Test cases covered : C150002 : Verify the login alert pop-up appear as per Store setting on supplemental roles and license whether they are enabled \ disabled
			 	     'C131088 : Verify user not able to login if user has license whose expiration date is blank

'Author : Kashish Ambwani

'Date Modified : 25 June 2017

'................................................................................................................................

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

WshSysEnv ("WF") = "C150002"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Import variables from sheet

advancedayssupplemental = DataTable.Value("AdvanceWarning_SupplementalRoles","Global")
advancedaysuserlicense = DataTable.Value("AdvanceWarning_UserLicense","Global")

'Call actions here


RunAction "Action1 [Preconditions_C150002_ECC]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Preconditions_C150002_EPS]", oneIteration
Wait(Iteration_Wait)
AdvanceWarningDays_SupplementalRoles(advancedayssupplemental)
Wait(Iteration_Wait)
AdvanceWarning_UserLicense(advancedaysuserlicense)
Wait(Iteration_Wait)
RunAction "Action1 [C150002_Steps]", oneIteration
Wait(Iteration_Wait)
AdvanceWarningDays_SupplementalRoles("")
Wait(Iteration_Wait)
AdvanceWarning_UserLicense("")
Wait(Iteration_Wait)
RunAction "Action1 [Revert_Changes_ECC]", oneIteration
Wait(Iteration_Wait)
PingECC() @@ hightlight id_;_17215934_;_script infofile_;_ZIP::ssf2.xml_;_


Reporter.ReportEvent micDone,"C150002","Test case has been executed"
