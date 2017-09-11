'..............................................................................................................................

'Test Name : ClientTimeout_SwitchUser_Main_RF

'Test Description : This test will peform the following steps :
					'1. Change Client inactive session timeout value to "0"
					'2. Logout from client.
					'3. Login to client and change client inactive session timeout to any value greater than "0"
					'4. Validate that Switch User is present.
					'5. Switch User and login with different User.
					'6. Logout from client.
					'7. Relaunch the client.
					'8. Wait until client inactive session timeout.
					'9. Click on relaunch button.
					'10. Again wait for client inactive session timeout.
					'11. Click on Exit.
					'12. Relaunch the client.
					'13. Change client inactive session timeout value to "0"
					
'Author : Kashish Ambwani

'Date Modified : 18th July 2017

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
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

WshSysEnv ("WF") = "SwitchUser"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Call actions here

CreateNewUserEPS()
Wait(Iteration_Wait)
RunAction "Action1 [ClientTimeout_SwitchUser_AT_RF]", oneIteration


Reporter.ReportEvent micDone,"Client Timeout and Switch User","Client Timeout and Switch User functionality has been tested" @@ hightlight id_;_5190150_;_script infofile_;_ZIP::ssf4.xml_;_
