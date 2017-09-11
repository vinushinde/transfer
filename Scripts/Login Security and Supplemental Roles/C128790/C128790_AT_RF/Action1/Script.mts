'........................................................................................................................................

'Test Name : Verify that user can change the password for EPS user from EPS client when user authentication mode is 2 (ECC Authentication).

'Test Description : Verify user can login with the changed password in mode2

'Test Rail ID : C128790

'Author : Kashish Ambwani

'Date Modified : 5th July 2017

'........................................................................................................................................


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

vurl="https://"&serverip&":58442/ecc/login.jsp"

WshSysEnv ("WF") = "C128790"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Import variables from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")
authenticationmode1 = DataTable.Value("AuthenticationMode","Global")
oldpassword2 = DataTable.Value("UserPassword","Global")
newpassword2 = DataTable.Value("UserNewPassword","Global")

newuser2 = RandomPtFirstName(5)
username2 = CStr(newuser2)

newuserinitials2 = RandomUserType(3)
userinitials2 = CStr(newuserinitials2)


user2group = DataTable.Value("UserGroup","Global")

user2licensetype = DataTable.Value("UserLicenseType","Global")

'Call actions

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AddNewUserECC_New username2,username2,username2,userinitials2,user2group,oldpassword2,user2licensetype
Wait(Iteration_Wait)
AddGroupUserECC(username2)
Wait(Iteration_Wait)
AuthenticationModeEPS(authenticationmode1)
Wait(Iteration_Wait)
AddLicenseUserEPS(username2)
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSChangePassword username2,newpassword2,oldpassword2
Wait(Iteration_Wait)
EPSLogin username2,oldpassword2
Wait(Iteration_Wait)
IncorrectAuthenticationEPS()
Wait(Iteration_Wait)
EPSLogin username2,newpassword2
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username,password
Wait(Iteration_Wait)
ValidateDeactivatedPassword(username2)
Wait(Iteration_Wait)
ECCLogout()
Wait(Iteration_Wait)
ECCLogin username2,oldpassword2
Wait(Iteration_Wait)
IncorrectAuthenticationECC()
Wait(Iteration_Wait)
ECCLogin username2,newpassword2
Wait(Iteration_Wait)
CloseIE()

Reporter.ReportEvent micDone,"C128790","Test Case has been exected successfully"
