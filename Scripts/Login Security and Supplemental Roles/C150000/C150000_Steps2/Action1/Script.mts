'..................................................................................

'Test Name : C150000_Steps2

'..................................................................................

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

'Import values from sheet

username2 = DataTable.Value("User2LogonID","Global")

'Click on Logout

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Enter User Auth.

JavaWindow("Enterprise Pharmacy System").JavaEdit("Username").Set DataTable.Value("User2LogonID","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("User2Pasword","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("Unable to login(st)").Check CheckPoint("Unable to login(st)") @@ hightlight id_;_13126322_;_script infofile_;_ZIP::ssf5.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaEdit("Username").Set DataTable.Value("UserName1","Global")

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Click
End If


'Query to find daily successful logins of user

SuccessfulLogins = "select ("+ vSchema +"USERS.DAILY_SUCCESSFUL_LOGONS) LOGINS from "+ vSchema +"USERS where "+ vSchema +"USERS.LOGIN = '"+ username2 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(SuccessfulLogins)

'assigning value to daily successful logins

logins = rcEPSRecordSet.Fields("LOGINS")


If logins = "0" Then
	Reporter.ReportEvent micPass,"Daily Successful Logons","0"
	Else
	Reporter.ReportEvent micFail,"Daily Successful Logons",""+ logins +""
End If

Reporter.ReportEvent micDone,"Steps","Steps have been performed"
