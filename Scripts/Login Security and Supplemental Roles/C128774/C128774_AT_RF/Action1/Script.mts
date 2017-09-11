'............................................................................................................................

'Test Name : Verify the EPS login- logout with EPS user when User Authentication mode is 2 (ECC Authentication) where ECC is up and running.

'Test Description : This test validates the ability of the user to login and logout of the system using Biometrics and User Name / Password Ldap up and down: Login/Logout System

'TestRail ID : C128774

'Author : Kashish Ambwani

'Date Modified : 13 July 2017

'............................................................................................................................

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

WshSysEnv ("WF") = "C128774"

vurl="https://"&serverip&":58442/ecc/login.jsp"

'Import data sheet

importDataSheet()

'Import variables from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")
eccauthentication = DataTable.Value("AuthenticationModeECC","Global")
epsauthentication = DataTable.Value("AuthenticationMode","Global")
password2 = DataTable.Value("Password2","Global")

'Call actions here

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AuthenticationModeECC(eccauthentication)
Wait(Iteration_Wait)
AuthenticationModeEPS(epsauthentication)
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username,password2
Wait(Iteration_Wait)
IncorrectAuthenticationEPS()
Wait(Iteration_Wait)
AuditUserAccessDB()
Wait(Iteration_Wait)
DeactivateUserEPSDB(username)
Wait(Iteration_Wait)
EPSLogin username,password
Wait(Iteration_Wait)
ValidateUserinactive()
Wait(Iteration_Wait)
AuditUserAccessDB()
Wait(Iteration_Wait)
ActivateUserEPSDB(username)
Wait(Iteration_Wait)
EPSLogin username,password
Wait(Iteration_Wait)


JavaWindow("Enterprise Pharmacy System").JavaObject("ExpandedWelcomeScreen").Check CheckPoint("ExpandedWelcomeScreen") @@ hightlight id_;_-1_;_script infofile_;_ZIP::ssf28.xml_;_

Wait(Iteration_Wait)

AuthenticationModeECC(epsauthentication)
Wait(Iteration_Wait)
CloseIECertificate()

Reporter.ReportEvent micDone,"C128774","Test case has been executed"
