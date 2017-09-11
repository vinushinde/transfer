'.....................................................................................................................................

'Test Name : C128786_Preconditions

'Test Description : This test case will setup the preconditions

'Author : Kashish Ambwani

'.....................................................................................................................................

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

'Import data sheet

DataTable.ImportSheet "E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\Corporate Authentication\C128786\C128786.xls",1,"Global"

'Import variables from sheet

username = DataTable.Value("Username","Global")
password = DataTable.Value("Password2","Global")
userauthenticationmode = DataTable.Value("AuthenticationMode","Global")
username2 = DataTable.Value("Login","Global")
password2 = DataTable.Value("Password","Global")


'Call actions here

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AuthenticationModeECC(userauthenticationmode)
Wait(Iteration_Wait)
ECCLogout()
Wait(Iteration_Wait)
ECCLogin username2,password2
Wait(Iteration_Wait)
AuthenticationModeEPS(userauthenticationmode)
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username2,password
Wait(Iteration_Wait)
IncorrectAuthenticationEPS()
Wait(Iteration_Wait)
AuditUserAccessDB()
Wait(Iteration_Wait)
EPSLogin username2,password2

Reporter.ReportEvent micDone,"Preconditions","All preconditions have been setup"
