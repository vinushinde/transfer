'.....................................................................................................................................

'Test Name : C128786_RevertChanges

'Test Description : This test case will revert all the changes made during the test case

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

authenticationmodenew = DataTable.Value("AuthenticationMode2","Global")
newusername = DataTable.Value("Username","Global")
newpassword = DataTable.Value("Password2","Global")


'Call actions here

AuthenticationModeEPS(authenticationmodenew)
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin newusername,newpassword
Wait(Iteration_Wait)
AuthenticationModeECC(authenticationmodenew)
Wait(Iteration_Wait)
CloseIECertificate()


Reporter.ReportEvent micDone,"Revert Changes","All changes have been reverted"
