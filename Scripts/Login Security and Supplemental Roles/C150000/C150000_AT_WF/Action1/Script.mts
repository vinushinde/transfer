'...............................................................................................................................................

'Test Name : Verify that user get login alert Halt as when the license expired or not added which are required to login as per state requirement.

'Test Description : Verify that user get login alert Halt as when the license expired or not added which are required to login as per state requirement.

'JIRA ID : C150000

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

WshSysEnv ("WF") = "C150000"

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data sheet

importDataSheet()

'Call variables from sheet

username1 = DataTable.Value("UserName","Global")
password1 = DataTable.Value("Password","Global")
licensename = DataTable.Value("LicenseTypeName","Global")
username2 = DataTable.Value("User2LogonID","Global")
password2 = DataTable.Value("User2Pasword","Global")


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
AddNewLicenseType_ECC(licensename)
Wait(Iteration_Wait)
RunAction "Action1 [AddNewUser1]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [AddNewLicenseState]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [AddNewUserECC]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [C150000_Steps2]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [AddNewUser2]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [C150000_Steps3]", oneIteration
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")
Wait(Iteration_Wait)
SupplementalRoleExpiration("Disabled")
Wait(Iteration_Wait)
RunAction "Action1 [Revert_Changes]", oneIteration

Reporter.ReportEvent micDone,"C150000","Test Case has been executed successfully"
