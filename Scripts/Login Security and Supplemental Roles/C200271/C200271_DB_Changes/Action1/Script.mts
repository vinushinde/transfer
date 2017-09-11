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

login2 = DataTable.Value("User_Login","Global")
strlogin2 = CStr(login2)
supprole2 = DataTable.Value("SupplementalRole1","Global")
strsupprole2 = CStr(supprole2)
strlicensenumber = CStr(DataTable.Value("UserLicenseNumber","Global"))


'Query to fetch user ID

userid = "select ("+ vSchema +"USERS.ID) USERID from "+ vSchema +"USERS where "+ vSchema +"USERS.LOGIN = '"+ strlogin2 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(userid)

'assigning value to user employee number

UserID = rcEPSRecordSet.Fields("USERID")
struserid = CStr(UserID)

Wait(Iteration_Wait)

'Query to fetch role ID of supplemental role

roleid = "select ("+ vSchema +"ROLES.ID) ROLEID from "+ vSchema +"ROLES where "+ vSchema +"ROLES.DESCRIPTION = '"+ strsupprole2 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(roleid)

'assigning value to role ID

RoleID = rcEPSRecordSet.Fields("ROLEID")
strroleid = CStr(RoleID)

Wait(Iteration_Wait)

'Query to update ending date on supplemental role

endsupprole = "update EPS2.SUPPLEMENTAL_ROLE_LINK set EPS2.SUPPLEMENTAL_ROLE_LINK.ENDING_DATE = sysdate+2 where EPS2.SUPPLEMENTAL_ROLE_LINK.ROLE_ID = '"+ strroleid +"' and EPS2.SUPPLEMENTAL_ROLE_LINK.USER_ID = '"+ struserid +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(endsupprole)

Wait(Iteration_Wait)

'Query to update deactivate date on license

deaclicense = "update EPS2.LICENSES set EPS2.LICENSES.DEACTIVATE_DATE = sysdate+2 where EPS2.LICENSES.LICENSE_OWNER_USER_ID = '"+ struserid +"' and EPS2.LICENSES.LICENSE_NUM = '"+ strlicensenumber +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(deaclicense)

Wait(Iteration_Wait)

'Query to update password change date on User

passchange = "update EPS2.USERS SET EPS2.USERS.PASSWORD_CHANGE_DATE = sysdate where EPS2.USERS.LOGIN = '"+ strlogin2 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(passchange)

Reporter.ReportEvent micDone,"DB Changes","All DB changes have been made"
