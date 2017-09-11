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

Wait(5)


'Query to find daily successful logins of user

SuccessfulLogins = "select ("+ vSchema +"USERS.DAILY_SUCCESSFUL_LOGONS) LOGINS from "+ vSchema +"USERS where "+ vSchema +"USERS.LOGIN = '"+ username2 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(SuccessfulLogins)

'assigning value to daily successful logins

logins = rcEPSRecordSet.Fields("LOGINS")
strlogins = CStr(logins)

If logins = "1" Then
	Reporter.ReportEvent micPass,"Daily Successful Logons","1"
	Else
	Reporter.ReportEvent micFail,"Daily Successful Logons",""+ strlogins +""
End If

Wait(2)

'Navigate to administration>User

JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_28617775_;_script infofile_;_ZIP::ssf10.xml_;_

'Search for user by employee ID

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1


JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set DataTable.Value("User2EmployeeID","Global")

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

Wait(2)

'Select first row for license numbers

JavaWindow("Enterprise Pharmacy System").JavaTable("State License Numbers").SelectRow "#0" @@ hightlight id_;_9615600_;_script infofile_;_ZIP::ssf13.xml_;_
Wait(1)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Edit License").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set "05-21-2017" @@ hightlight id_;_15502042_;_script infofile_;_ZIP::ssf16.xml_;_


If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaTable("State License Numbers").SelectRow "#1" @@ hightlight id_;_9615600_;_script infofile_;_ZIP::ssf18.xml_;_
If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Edit License").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set "05-21-2017" @@ hightlight id_;_15502042_;_script infofile_;_ZIP::ssf16.xml_;_


If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

'Enter User Authentication


JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("User2LogonID","Global") @@ hightlight id_;_20085272_;_script infofile_;_ZIP::ssf23.xml_;_
Wait(1)
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("User2Pasword","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

'Click on Logout

JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").WaitProperty "visible",1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Enter User Auth.

JavaWindow("Enterprise Pharmacy System").JavaEdit("Username").Set DataTable.Value("User2LogonID","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("User2Pasword","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Click
End If @@ hightlight id_;_20705975_;_script infofile_;_ZIP::ssf30.xml_;_

'Checkpoint for Unable to Login alert

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("Unable to login(st)").Check CheckPoint("Unable to login(st)_2") @@ hightlight id_;_5529715_;_script infofile_;_ZIP::ssf31.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If

'Enter user authentication with akumar


JavaWindow("Enterprise Pharmacy System").JavaEdit("Username").Set DataTable.Value("UserName","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Clsoe Internet Browser @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center")_;_script infofile_;_ZIP::ssf36.xml_;_

If Browser("Enterprise Control Center").Exist(15) Then
Browser("Enterprise Control Center").Close
End If


Reporter.ReportEvent micDone,"Steps","Steps have been performed"
