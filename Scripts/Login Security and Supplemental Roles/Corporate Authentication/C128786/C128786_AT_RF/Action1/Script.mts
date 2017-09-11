'........................................................................................................................................

'Test Name : C128786_AT_RF

'........................................................................................................................................


'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

'Call actions here


RunAction "Action1 [CreateUser_AD_C128786]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [AD_Setup_Main_RF]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [C128786_Preconditions]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [C128786_RevertChanges]", oneIteration


Reporter.ReportEvent micDone,"C128786","Test case has been executed successfully"

