'........................................................................................................................

'Test Name : ClientTimeout_SwitchUser_AT_RF

'........................................................................................................................


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
	
	WF_Mode = WshSysEnv("vMODE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword
vServerip = WshSysEnv("vServerip")

'Import variables from sheet

clienttimeoutzero = DataTable.Value("ClientTimeout","Global")
clienttimeoutlogout = DataTable.Value("ClientTimeout_Logout","Global")
username1 = DataTable.Value("Username","Global")
password1 = DataTable.Value("Password","Global")
username2 = DataTable.Value("Username2","Global")
password2 = DataTable.Value("Password2","Global")

clienttimeoutwait = clienttimeoutlogout*60

'...........................................................................................................................................

'Call steps

ClientTimeout(clienttimeoutzero)
Wait(Iteration_Wait)

'...........................................................................................................................................

'Checkpoint that Switch User Button should not be displayed

If JavaWindow("Enterprise Pharmacy System").JavaButton("Switch User").Exist(15) Then
	Reporter.ReportEvent micFail,"Switch User button when client inactive session timeout is zero","Displayed"
	Else
	Reporter.ReportEvent micPass,"Switch User button when client inactive session timeout is zero","Not Displayed"
End If

'...........................................................................................................................................

EPSLogout()
Wait(Iteration_Wait)

'...........................................................................................................................................

'Checkpoint for Login Screen

'Check if Username and Password fields are visible

JavaWindow("Enterprise Pharmacy System").JavaStaticText("Username(st)").Check CheckPoint("Username(st)") @@ hightlight id_;_19061918_;_script infofile_;_ZIP::ssf8.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaStaticText("Password(st)").Check CheckPoint("Password(st)") @@ hightlight id_;_24294079_;_script infofile_;_ZIP::ssf9.xml_;_

'...........................................................................................................................................

EPSLogin username1,password1
Wait(Iteration_Wait)
ClientTimeout(clienttimeoutlogout)
Wait(Iteration_Wait)

'...........................................................................................................................................

'Checkpoint for Switch User button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Switch User").GetROProperty("visible")=1 Then
	Reporter.ReportEvent micPass,"Switch User button when client inactive session timeout is greater than zero","Displayed"
	Else
	Reporter.ReportEvent micFail,"Switch User button when client inactive session timeout is greater than zero","Not Displayed"
End If

'...........................................................................................................................................
 @@ hightlight id_;_31515361_;_script infofile_;_ZIP::ssf14.xml_;_
 'Click on Switch User
 
 If JavaWindow("Enterprise Pharmacy System").JavaButton("Switch User").Exist(15) Then
 	JavaWindow("Enterprise Pharmacy System").JavaButton("Switch User").Click
 End If

'Login with another User
 
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set username2

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set password2

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

 
'Checkpoint to see if User has been switched

If WF_Mode = "RAPIDFILL" Then
JavaWindow("Enterprise Pharmacy System").JavaObject("ExpandedWelcomeScreen").Check CheckPoint("ExpandedWelcomeScreen")
ElseIf WF_Mode = "WORKFLOW" Then
JavaWindow("Enterprise Pharmacy System").JavaObject("ExpandedWelcomeScreen").Check CheckPoint("ExpandedWelcomeScreen_2") @@ hightlight id_;_-1_;_script infofile_;_ZIP::ssf33.xml_;_
End If


'...........................................................................................................................................
 
 EPSLogout()
 Wait(Iteration_Wait)
 
'...........................................................................................................................................
 
 'Relaunch client @@ hightlight id_;_28026180_;_script infofile_;_ZIP::ssf17.xml_;_
 
 Call LaunchEPS_Sprint(username1,password1,vServerip)

 Wait(clienttimeoutwait)

'.............................................................................................................................................

'Checkpoint for Session has ended due to inactivity pop-up

JavaDialog("Session Timeout").JavaStaticText("Session has ended due").Check CheckPoint("Session has ended due to inactivity.(st)") @@ hightlight id_;_32363551_;_script infofile_;_ZIP::ssf19.xml_;_

JavaDialog("Session Timeout").JavaButton("Relaunch").Check CheckPoint("Relaunch") @@ hightlight id_;_18731070_;_script infofile_;_ZIP::ssf20.xml_;_

JavaDialog("Session Timeout").JavaButton("Exit").Check CheckPoint("Exit") @@ hightlight id_;_13268106_;_script infofile_;_ZIP::ssf21.xml_;_

'Click on Re-Launch

If JavaDialog("Session Timeout").JavaButton("Relaunch").Exist(15) Then
JavaDialog("Session Timeout").JavaButton("Relaunch").Click
End If

Wait(20)

	If JavaDialog("Security Information").Exist(30) Then
	JavaDialog("Security Information").JavaButton("Run").Click
	End If
	
	If JavaDialog("Security Warning").Exist(10) Then
	JavaDialog("Security Warning").JavaButton("Continue").Click
	End If
	
	If JavaDialog("Security Information").Exist(10) Then
	JavaDialog("Security Information").JavaButton("Run").Click
	End If
	
	If JavaDialog("Security Warning").Exist(10) Then
	JavaDialog("Security Warning").JavaButton("Continue").Click
	End If
	
	If JavaDialog("Security Information").Exist(20) Then
	JavaDialog("Security Information").JavaButton("Run").Click
	End If
	
	Wait(20)

'.............................................................................................................................................


JavaDialog("Enterprise Pharmacy System").JavaEdit("Username").Set username1

JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set password1

If JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

Wait(10)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Expired Data Files").Exist(50) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Expired Data Files").JavaButton("OK").Click
End If


Wait(clienttimeoutwait)

'.............................................................................................................................................

'Checkpoint for Session has ended due to inactivity pop-up

JavaDialog("Session Timeout").JavaStaticText("Session has ended due").Check CheckPoint("Session has ended due to inactivity.(st)") @@ hightlight id_;_32363551_;_script infofile_;_ZIP::ssf19.xml_;_

JavaDialog("Session Timeout").JavaButton("Relaunch").Check CheckPoint("Relaunch") @@ hightlight id_;_18731070_;_script infofile_;_ZIP::ssf20.xml_;_

JavaDialog("Session Timeout").JavaButton("Exit").Check CheckPoint("Exit")

'Click on Exit

If JavaDialog("Session Timeout").Exist(15) Then
JavaDialog("Session Timeout").JavaButton("Exit").Click
End If

'.............................................................................................................................................

'Relaunch Client

 Call LaunchEPS_Sprint(username1,password1,vServerip)

'Change client innactive session timeout to zero

ClientTimeout(clienttimeoutzero)


Reporter.ReportEvent micDone,"Steps","All steps have been executed"
