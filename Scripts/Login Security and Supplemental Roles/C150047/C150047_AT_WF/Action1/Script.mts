'...............................................................................................................................................

'Test Name : Verify that the ECC user can add/activate/deactivate a custom License types and can be assigned to a user license .

'Test Description : Verify that the ECC user can add/activate/deactivate a custom License types and can be assigned to a user license.

'JIRA ID : C150047

'Author : Kashish Ambwani

'Date Modified : 19 June 2017

'...............................................................................................................................................

'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

WshSysEnv ("WF") = "C150047"

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data sheet

importDataSheet()

'Call variables from sheet

username1 = DataTable.Value("UserName","Global")
password1 = DataTable.Value("Password","Global")
userpermission = DataTable.Value("UserPermission","Global")
authentication = DataTable.Value("AuthenticationMode","Global")

'Call actions here

AuthenticationModeEPS(authentication)
Wait(Iteration_Wait)
LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username1,password1
Wait(Iteration_Wait)
AddRoleECCUser(userpermission)
Wait(Iteration_Wait)
ECCLogin username1,password1
Wait(Iteration_Wait)
RunAction "Action1 [C150047_Steps]", oneIteration
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")


Reporter.ReportEvent micDone,"C150047","Test Case has been executed successfully" @@ hightlight id_;_33121334_;_script infofile_;_ZIP::ssf7.xml_;_
