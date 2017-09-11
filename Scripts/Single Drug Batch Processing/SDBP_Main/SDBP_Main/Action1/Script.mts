'......................................................................................................................................

'Test Cases covered : C131184 : Validate when selecting SEL rel code all fields on granted by are automatically populated (RF/WF Mode)
					 'C131166 : Single Drug Batch Processing: Verify the batch review screen, Site of Administration selection and processing, Pharmacist Verification and its contents for single drug batch processing.
					 'C131172 : WF : "GetNext" ability for these transactions is not allowed for SDBP Rx's
					 'C131175 : Do not Run RxEdit checks for the batch records but run the DURs.
					 'C131177 : Single Drug Batch Processing - Add and verify Incentive Fees for TP transactions with multiple Tp's
					 'C131179 : Batch Review from Open Orders screen
					 'C131180 : COB Review from Open Orders screen

'Author : Kashish Ambwani

'Date Modified : 7 August 2017

'......................................................................................................................................
 @@ hightlight id_;_20590695_;_script infofile_;_ZIP::ssf181.xml_;_
 
	'Regression testing connection:
	Set WshShell = CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("User")
	Set EPS2DBObject = CreateObject("ADODB.Connection")
	Dim    vSchema 	, vEnvironment, vDSN
	vSchema  =  WshSysEnv("epsvSchema")
	vEnvironment = WshSysEnv("epsEnvironment")
	vDSN               =  WshSysEnv("vDSN")	
	dbpassword = WshSysEnv("vDbpwd")
	dbuser = WshSysEnv("vDbuser")
	serverip = WshSysEnv("vServerip")
	WF_mode = WshSysEnv("vMODE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

WshSysEnv ("WF") = "SDBPMain"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Call actions here

RunAction "Action1 [Test_Data_SDBP]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Preconditions] [2]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Create_New_Batch_New]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [SDBP_Open_Orders_New]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [TPException_Queue_New]", oneIteration
Wait(Iteration_Wait)
If WF_mode = "WORKFLOW" Then
RunAction "Action1 [SDBP_Open_Orders_COB]", oneIteration
Wait(Iteration_Wait)
End If
RunAction "Action1 [TP_Transmit_Queue_New]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Validations]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [SDBP_Pharmacist_Verification_New]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Revert Changes] [2]", oneIteration

Reporter.ReportEvent micDone,"Single Drug Batch Processing","Single Drug Batch Processing has been tested"
