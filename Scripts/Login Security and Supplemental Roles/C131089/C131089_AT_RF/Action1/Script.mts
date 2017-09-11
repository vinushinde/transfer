'.................................................................................................................................

'Test Name : Verify user not able to login if user has license whose expiration date is blank and check the ecc authentication

'Test Description : Test ECC Authentication at WF store from ECC while EPS server is up.

'Test Rail ID : C131089

'Date Modified : 12 July 2017

'Author : Kashish Ambwani

'.................................................................................................................................

'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

WshSysEnv ("WF") = "C131089"

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import sheet

importDataSheet()

'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")
authenticationmode1 = DataTable.Value("AuthenticationMode","Global")
username2 = DataTable.Value("User_Login","Global")
password2 = DataTable.Value("User_Password","Global")


'Call actions here

AuthenticationModeEPS(authenticationmode1)
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")
Wait(Iteration_Wait)
LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AuthenticationModeECC(authenticationmode1)
Wait(Iteration_Wait)
CloseIECertificate()
Wait(Iteration_Wait)
PingECC()
Wait(Iteration_Wait)
CreateNewUserEPS()
Wait(Iteration_Wait)
AddLicenseUserEPS(username2)
Wait(Iteration_Wait)
UserLicesnseExpiration("Enabled")
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username2,password2
Wait(Iteration_Wait)
ValidateUnabletoLogin()
Wait(Iteration_Wait)
EPSLogin username,password
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")

Reporter.ReportEvent micDone,"C131089","Test case has been exected successfully" @@ hightlight id_;_26895097_;_script infofile_;_ZIP::ssf50.xml_;_
