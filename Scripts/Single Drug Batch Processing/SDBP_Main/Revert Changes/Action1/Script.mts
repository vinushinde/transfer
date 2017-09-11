'..............................................................................................................................

'Test Name : This test will revert all the changes made in the preconditions

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
JavaWindow("Enterprise Pharmacy System").Restore @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf49.xml_;_
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


'Set single drug batch processing to No

SDBPEnabled("No")

Wait(Min_Wait)

If releasebase = "2608" Then
SDBPApplicationSettings_2608 "No","No","No","No","No","No","No","No","No"
ElseIf releasebase = "2609" Then
SDBPApplicationSettings_2609 "No","No","No","No","No","No","No","No","No","No","No"
End If

'Add incentive fee in application settings

DefaultIncentiveFee("")

Wait(Min_Wait)

'TP1 set to pay

prescriberID tp1,prescriberpaid

Wait(Min_Wait)

'Split bill review should be unchecked on all 3 TPs

TP_SplitBillReview tp1,"OFF" @@ hightlight id_;_26875026_;_script infofile_;_ZIP::ssf26.xml_;_
Wait(2)
TP_SplitBillReview tp2,"OFF"
Wait(2)
TP_SplitBillReview tp3,"OFF"

Wait(Min_Wait)

'Set Patient Birth Order to null

PatBirthOrder "#0","#0"

Reporter.ReportEvent micDone,"Revert Changes","All changes have been reverted" @@ hightlight id_;_8393000_;_script infofile_;_ZIP::ssf48.xml_;_
