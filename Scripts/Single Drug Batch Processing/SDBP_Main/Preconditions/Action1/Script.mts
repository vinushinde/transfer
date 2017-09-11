'..............................................................................................................................

'Test Name : This test will set all the preconditions for the test case.

'..............................................................................................................................


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
	releasebase = WshSysEnv("vRELEASE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

'Import variables from sheet

incentivefee = DataTable.Value("IncentiveFee","Global")
prescriberpaid = DataTable.Value("Prescriber_Paid_ID","Global")
prescriberreject = DataTable.Value("Prescriber_Reject_ID","Global")
tp1 = DataTable.Value("ThirdParty1","Global")
tp2 = DataTable.Value("ThirdParty2","Global")
tp3 = DataTable.Value("ThirdParty3","Global")


'Query to put all tasks on hold

hold = "update "+ vSchema +"WORKFLOW_TOKEN set "+ vSchema +"WORKFLOW_TOKEN.HOLD_UNTIL_DATE= sysdate +5"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(hold)

'Set single drug batch processing to Yes

SDBPEnabled("Yes")

Wait(Min_Wait)

'Set all values in single drug batch processing on application settings

If releasebase = "2608" Then
SDBPApplicationSettings_2608 "Yes","Yes","Yes","Yes","Yes","Yes","Yes","Yes","Yes"
ElseIf releasebase = "2609" Then
SDBPApplicationSettings_2609 "Yes","Yes","Yes","Yes","No","Yes","Yes","Yes","Yes","No","No"
End If


Wait(Min_Wait)

'Check SDBP on drug

SDBPDrug("ON")

Wait(Min_Wait)

'Add incentive fee in application settings

DefaultIncentiveFee(incentivefee)

Wait(Min_Wait)

'TP1 set to reject and TP2 and TP3 set to pay

prescriberID tp1,prescriberreject
Wait(3)
prescriberID tp2,prescriberpaid
Wait(3)
prescriberID tp3,prescriberpaid

Wait(Min_Wait)

'COB should be checked on all 3 TPs

ThirdParty_COB_Checkbox tp1,"ON"
Wait(2)
ThirdParty_COB_Checkbox tp2,"ON"
Wait(2)
ThirdParty_COB_Checkbox tp3,"ON"

Wait(Min_Wait)

'Split bill review should be checked on all 3 TPs

TP_SplitBillReview tp1,"ON" @@ hightlight id_;_26875026_;_script infofile_;_ZIP::ssf26.xml_;_
Wait(2)
TP_SplitBillReview tp2,"ON"
Wait(2)
TP_SplitBillReview tp3,"ON"


Reporter.ReportEvent micDone,"Preconditions","All preconditions have been set"
