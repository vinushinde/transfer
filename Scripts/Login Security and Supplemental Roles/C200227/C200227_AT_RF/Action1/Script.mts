'.....................................................................................................................................

'Test Name : C200227_AT_RF

'.....................................................................................................................................

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

'Import variables from sheet

authenticationmode1 = DataTable.Value("AuthenticationMode1","Global")
authenticationmode2 = DataTable.Value("AuthenticationMode2","Global")
maximumpasswordduration = DataTable.Value("MaximumPasswordDuration","Global")
advancewarningexpiration = DataTable.Value("AdvanceWarningofExpiration","Global")
username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

'Call actions

passwordapplicationsettings maximumpasswordduration,advancewarningexpiration
Wait(Iteration_Wait)
AuthenticationModeEPS(authenticationmode1)
Wait(Iteration_Wait)

'Query to update the password change date on User
passwordchange = "update "+ vSchema +"USERS SET "+ vSchema +"USERS.PASSWORD_CHANGE_DATE = sysdate where "+ vSchema +"USERS.LOGIN = '"+ username +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(passwordchange)

Wait(Iteration_Wait)

'Logout of EPS

EPSLogout()

Wait(Iteration_Wait)

'Login to EPS

EPSLogin username,password

Wait(Iteration_Wait)

'Checkpoint for password reset pop-up

JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaStaticText("Your password will expire").Check CheckPoint("Your password will expire in 2 days. Reset your password.(st)_2") @@ hightlight id_;_32106880_;_script infofile_;_ZIP::ssf16.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaButton("OK").Click
End If
 @@ hightlight id_;_25950477_;_script infofile_;_ZIP::ssf10.xml_;_
 Wait(Iteration_Wait)
 
 'Change Authentication Mode to ECC Authentication
 
 AuthenticationModeEPS(authenticationmode2)
 
 Wait(Iteration_Wait)
 
 'Logout of EPS
 
 EPSLogout()
 
 Wait(Iteration_Wait)
 
 'Login to EPS
 
 EPSLogin username,password
 
 Wait(Iteration_Wait)
 
'Checkpoint for password reset pop-up

JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaStaticText("Your password will expire").Check CheckPoint("Your password will expire in 2 days. Reset your password.(st)_2") @@ hightlight id_;_32106880_;_script infofile_;_ZIP::ssf16.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaButton("OK").Click
End If

'Query to update the password change date on User
passwordchangereset = "update "+ vSchema +"USERS SET "+ vSchema +"USERS.PASSWORD_CHANGE_DATE = sysdate+180 where "+ vSchema +"USERS.LOGIN = '"+ username +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(passwordchangereset)

Wait(Iteration_Wait)

passwordapplicationsettings "365","3"

Reporter.ReportEvent micDone,"Test Steps","All the test steps have been executed"
