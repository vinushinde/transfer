'...............................................................................................................................................

'Test Name : Verify the functionality of activating/deactivating of license in EPS license requirement screen on ECC

'Test Description : Verify the functionality of activating/deactivating of license in EPS license requirement screen on ECC.

'JIRA ID : C150048

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

WshSysEnv ("WF") = "C150048"

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data sheet

importDataSheet()

'Call variables from sheet

username1 = DataTable.Value("UserName","Global")
password1 = DataTable.Value("Password","Global")
userpermission = DataTable.Value("UserPermission","Global")




'Call actions here

UserLicesnseExpiration("Enabled")
Wait(Iteration_Wait)
SupplementalRoleExpiration("Enabled")
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
RunAction "Action1 [C150048_Steps]", oneIteration
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")
Wait(Iteration_Wait)
SupplementalRoleExpiration("Disabled")



Reporter.ReportEvent micDone,"C150048","Test Case has been executed successfully" @@ hightlight id_;_Browser("Enterprise Control Center 2").Page("Enterprise Control Center 5").WebArea("https://192.168.109.121:58442/ecc/i")_;_script infofile_;_ZIP::ssf4.xml_;_
