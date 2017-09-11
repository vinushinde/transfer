'...........................................................................................................................

'Test Name : C200271_DB_Changes

'...........................................................................................................................

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

'Import data from table

login3 = DataTable.Value("User_Login","Global")
strlogin3 = CStr(login3)
rolename = DataTable.Value("SupplementalRole1","Global")


'Query to fetch user ID

userid2 = "select ("+ vSchema +"USERS.ID) USERID from "+ vSchema +"USERS where "+ vSchema +"USERS.LOGIN = '"+ strlogin3 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(userid2)

'assigning value to user employee number

UserID2 = rcEPSRecordSet.Fields("USERID")
struserid2 = CStr(UserID2)

Wait(Iteration_Wait)

'Query to update deactivate date on license

deaclicense2 = "update EPS2.LICENSES set EPS2.LICENSES.DEACTIVATE_DATE = sysdate-2 where EPS2.LICENSES.LICENSE_OWNER_USER_ID = '"+ struserid2 +"' and EPS2.LICENSES.LICENSE_NUM = '"+ strlicensenumber +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(deaclicense2)

Wait(Iteration_Wait)

'Query to update password change date on User

passchange2 = "update EPS2.USERS SET EPS2.USERS.PASSWORD_CHANGE_DATE = sysdate+365 where EPS2.USERS.LOGIN = '"+ strlogin3 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(passchange2)
 @@ hightlight id_;_Browser("MSN India | Hotmail, Outlook,").Page("Enterprise Control Center 2").WebElement("WebTable")_;_script infofile_;_ZIP::ssf11.xml_;_

Wait(Iteration_Wait) @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Control Center 2").WebElement("WebTable")_;_script infofile_;_ZIP::ssf22.xml_;_

RemoveSupplementalRole(rolename)

Wait(Iteration_Wait)

CloseIECertificate()

Wait(Iteration_Wait)

PingECC() @@ hightlight id_;_1778615_;_script infofile_;_ZIP::ssf24.xml_;_

Reporter.ReportEvent micDone,"Revert Changes","All changes have been reverted"
