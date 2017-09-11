'.....................................................................................................................................

'Test Name : AD_Setup_Main_RF

'Test Description : This test case will setup the communication and binding settings for Corporate Authentication on ECC

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

DataTable.ImportSheet "E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\Corporate Authentication\AD_SETUP\ADSetup.xls",1,"Global"

'Import variables from sheet

usernameecc = DataTable.Value("UsernameECC","Global")
passwordecc = DataTable.Value("PasswordECC","Global")

'Call actions here

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin usernameecc,passwordecc
Wait(Iteration_Wait)
ADCommSettings()
Wait(Iteration_Wait)
ADBindingSettings()
Wait(Iteration_Wait)
CloseIECertificate()


Reporter.ReportEvent micDone,"AD Setup","The communication and binding settings for corporate authentication have been setup on ECC"
