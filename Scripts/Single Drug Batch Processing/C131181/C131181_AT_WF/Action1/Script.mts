'................................................................................................................................

'Test Name : Cancel Batch Rx from Open Orders screen

'JIRA ID : C131181

'Author : Kashish Ambwani

'Date Modified : 9 June 2017

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
	releasebase = WshSysEnv("vRELEASE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword


WshSysEnv ("WF") = "C131181"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Import variables from sheet

multiplebirth = DataTable.Value("PatientMultipleBirth","Global")
userbirthorder = DataTable.Value("PatientBirthOrder","Global")
userpatrace = DataTable.Value("PatientRace","Global")

'Call actions here


RunAction "Action1 [Create_Patient]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Create_Prescriber]", oneIteration
Wait(Iteration_Wait)
SDBPEnabled("Yes")
Wait(Iteration_Wait)
If releasebase = "2608" Then
SDBPApplicationSettings_2608 "No","No","No","No","Yes","Yes","No","No","No"
ElseIf releasebase = "2609" Then
SDBPApplicationSettings_2609 "No","No","No","No","Yes","Yes","No","No","No","No","No"
End If
Wait(Iteration_Wait)
SDBPDrug("ON")
Wait(Iteration_Wait)
PatBirthOrder multiplebirth,userbirthorder
Wait(Iteration_Wait)
PatRace(userpatrace)
Wait(Iteration_Wait)
RunAction "Action1 [Create_New_Batch]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [SDBP_Open_Orders]", oneIteration
Wait(Iteration_Wait)
SDBPEnabled("No")
Wait(Iteration_Wait)
SDBPDrug("OFF")
Wait(Iteration_Wait)
If releasebase = "2608" Then
SDBPApplicationSettings_2608 "No","No","No","No","No","No","No","No","No"
ElseIf releasebase = "2609" Then
SDBPApplicationSettings_2609 "No","No","No","No","No","No","No","No","No","No","No"
End If

Reporter.ReportEvent micDone,"C131181","Test case has been executed successfully"
