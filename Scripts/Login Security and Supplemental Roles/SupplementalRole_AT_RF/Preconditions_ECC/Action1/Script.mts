'...............................................................................................................................................

'Test Name : ECC_Part

'Date Modified : 24 June 2017
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

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")
userpermisssion = DataTable.Value("UserPermission","Global")

'Call actions here

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
RemoveRoleECCUser(userpermisssion)
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AddRoleECCUser (userpermisssion)
Wait(Iteration_Wait)
ECCLogin username,password


Reporter.ReportEvent micDone,"Preconditions ECC","All preconditions on ECC have been set" @@ hightlight id_;_Browser("Enterprise Control Center 2").Page("Enterprise Control Center 4").WebElement("WebTable")_;_script infofile_;_ZIP::ssf51.xml_;_
