'.......................................................................................................................................

'Test Name : Validations.

'Test Description : This test will cover all the valdiations of the test case.

'.......................................................................................................................................


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
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword


'Import data from sheet into variables

Dim strfirstname,strlastname

strfirstname = DataTable.Value("PT_FIRSTNAME","Global")
strlastname = DataTable.Value("PT_LASTNAME","Global")	

'Query to fetch the highest Rx Number  from table
RxNumber = "Select max("+ vSchema +"RX_SUMMARY.RX_NUMBER) RXNUM from "+ vSchema +"RX_SUMMARY,"+ vSchema +"PATIENT where "+ vSchema +"RX_SUMMARY.PATIENT_ID = "+ vSchema +"PATIENT.ID and "+ vSchema +"PATIENT.LAST_NAME = '"+ strlastname +"' and "+ vSchema +"PATIENT.FIRST_NAME = '"+ strfirstname +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(RxNumber)

'assigning value to rxnumber
rxnumber = rcEPSRecordSet.Fields("RXNUM")
strrxnumber = CStr(rxnumber)
Wait(3)

'Query to fetch rx_tx_id from RX_TX table

RxTxID = "select ("+ vSchema +"RX_TX.ID) RXTXID from "+ vSchema +"RX_TX,"+ vSchema +"RX_SUMMARY where "+ vSchema +"RX_TX.RX_SUMMARY_ID = "+ vSchema +"RX_SUMMARY.ID and "+ vSchema +"RX_SUMMARY.RX_NUMBER = '"+ strrxnumber +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(RxTxID)

'assigning value to rxnumber
rxtxid = rcEPSRecordSet.Fields("RXTXID")
strrxtxid = CStr(rxtxid)
Wait(3)

'Query to fetch rx/tx values from DB
RxTxRelation = "select ("+ vSchema +"RX_TX.CONSENT_BY_RELATION_CD) RELATION ,("+ vSchema +"RX_TX.CONSENT_BY_FIRST_NAME) FNAME,("+ vSchema +"RX_TX.CONSENT_BY_LAST_NAME) LNAME,("+ vSchema +"RX_TX.CONSENT_BY_MIDDLE_NAME) MNAME from "+ vSchema +"RX_TX where "+ vSchema +"RX_TX.ID = '"+ strrxtxid +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(RxTxRelation)

'assigning value to rxtx fields

rxtxrelation = rcEPSRecordSet.Fields("RELATION")
rxtxmname = rcEPSRecordSet.Fields("MNAME")
rxtxlname = rcEPSRecordSet.Fields("LNAME")
rxtxfname = rcEPSRecordSet.Fields("FNAME")

datarelation = CStr(DataTable.Value("CONSENT_GRANTED_BY_RELATION","Global"))
datafirstname = CStr(DataTable.Value("PT_FIRSTNAME","Global"))
datalastname = CStr(DataTable.Value("PT_LASTNAME","Global"))
datamiddlename = CStr(DataTable.Value("PT_MIDDLENAME","Global"))

If datarelation = rxtxrelation and datafirstname = rxtxfname and datalastname = rxtxlname and datamiddlename = rxtxmname Then
	
	Reporter.ReportEvent micPass,"DB Validation - Consent By Relation","All values are correct"
	Else
	Reporter.ReportEvent micFail,"DB Validation - Consent By Relation","Incorrect"
	
End If

Wait(2)

'Query to find if Rx_Edits have run

RxEdits = "select count("+ vSchema +"RX_EDIT_RESULTS.ID) RXEDITS from "+ vSchema +"RX_EDIT_RESULTS where "+ vSchema +"RX_EDIT_RESULTS.RX_TX_ID = '"+ strrxtxid +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(RxEdits)

'assigning value to rx edits
rxedits = rcEPSRecordSet.Fields("RXEDITS")


If rxedits = "0" Then
	Reporter.ReportEvent micPass,"Rx Edits","No rx Edits have run"
	Else
	Reporter.ReportEvent micFail,"Rx Edits","Rx Edits have run"
End If

Wait(2)


'JavaWindow("Enterprise Pharmacy System").JavaTable("Tasks for This Transaction").SelectRow "#8"

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next Task").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next Task").Click
End If

'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global")
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

If JavaDialog("No tasks").Exist(15) Then

Reporter.ReportEvent micPass,"Get Next - SDBP Rx(s)","User is not able to get Single Drug Batch Processing Rx(s) from Get Next"

JavaDialog("No tasks").JavaButton("OK").Click
End If

'Navigate to Filecabinet>Patient>Rx/Tx

JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Patient").JavaMenu("Rx/Tx").Select

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").Set DataTable.Value("PT_LASTNAME","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("First Name").Set DataTable.Value("PT_FIRSTNAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

'Set Rx number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx or Tx Number").WaitProperty "visible",1

'Filter Rx by Batch Filled Rx option from dropdown

JavaWindow("Enterprise Pharmacy System").JavaList("Filter by").Select "Batch Filled Rx"

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx or Tx Number").Set strrxnumber

If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
End If

'Validation for batch filled Rx on EPS

'Query to fetch the highest Rx Number  from table
batchfilled = "select ("+ vSchema +"LINE_ITEM.LINE_ITEM_TYPE) ITEMTYPE from "+ vSchema +"LINE_ITEM,"+ vSchema +"RX_SUMMARY,"+ vSchema +"RX_TX where "+ vSchema +"LINE_ITEM.ID = "+ vSchema +"RX_TX.ID and "+ vSchema +"RX_TX.RX_SUMMARY_ID = "+ vSchema +"RX_SUMMARY.ID and "+ vSchema +"RX_SUMMARY.RX_NUMBER = '"+ strrxnumber +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(batchfilled)

'assigning value to line item type
itemtype = rcEPSRecordSet.Fields("ITEMTYPE")

If itemtype = "18" Then
	Reporter.ReportEvent micPass,"DB Validation - Line Item Type","18 - Batch Filled Rx"
	Else
	Reporter.ReportEvent micFail,"DB Validation - Line Item Type",""+ itemtype +""
End If


If JavaWindow("Enterprise Pharmacy System").JavaButton("Tx Detail").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Tx Detail").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Task Tracking").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Task Tracking").Click
End If

'Check if DUR has run for Batch processed Rx

getrows = JavaWindow("Enterprise Pharmacy System").JavaTable("Tasks for This Transaction").GetROProperty("rows")
	
For i = 0 To getrows-1 Step 1
	
	gettaskdescription = JavaWindow("Enterprise Pharmacy System").JavaTable("Tasks for This Transaction").GetCellData(i,"Task Description")
	getaction = JavaWindow("Enterprise Pharmacy System").JavaTable("Tasks for This Transaction").GetCellData(i,"Action")
	
If gettaskdescription = "Drug Utilization Review" and getaction = "complete" Then
		Reporter.ReportEvent micPass,"SDBP - DUR","DUR has run for Batch Processed Rx"
	Exit For
End If

Next


If JavaWindow("Enterprise Pharmacy System").JavaButton("Close Tx Detail").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Close Tx Detail").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Reporter.ReportEvent micDone,"validations","All validations have been completed"
