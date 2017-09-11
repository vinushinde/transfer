'....................................................................................................................................

'Test Cases Covered : C150046 : Verify that the user is able to ADD / REMOVE the supplemental roles from ECC > administration > user security > role setting
					 'C150045 : Verify that user can add new and view expired supplemental roles to a User on EPS > Admin > User Screen and is able to access the features which have a supplemental role defined.
					 

'Author : Kashish Ambwani

'Date Modified : 24 June 2017

'....................................................................................................................................


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

DataTable.ImportSheet "E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\SupplementalRole_AT_RF\SupplementalRole.xls",1,"Global"

'Import data into variables from sheet

authenticationmode = DataTable.Value("AuthenticationMode","Global")


'Call Actions here


RunAction "Action1 [Preconditions_ECC]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Validations_ECC]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Preconditions_ECC2]", oneIteration
Wait(Iteration_Wait)
AuthenticationModeEPS(authenticationmode) @@ hightlight id_;_33121334_;_script infofile_;_ZIP::ssf8.xml_;_
Wait(Iteration_Wait)
RunAction "Action1 [EPS_Steps]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [RevertChanges_ECC]", oneIteration


Reporter.ReportEvent micDone,"Supplemental Role","Test case has been executed"



