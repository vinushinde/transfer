'..................................................................................................................................

'Test Name : Create_New_Batch

'Test Description : This test will create a new batch for single drug batch processing

'Author : Kashish Ambwani

'..................................................................................................................................

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


WF_Value = WshSysEnv ("WF")

'Navigate to Tools>Utilities>Single Drug Batch Processing>Create New Batch
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Create New Batch").Select @@ hightlight id_;_17891018_;_script infofile_;_ZIP::ssf2.xml_;_

'Enter all mandatory details on Drug Selection Screen

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").WaitProperty "visible",1

'Enter Fill Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Fill Date").Set DataTable.Value("Fill_Date","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Fill Date").Type micTab

'Enter DAW

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").Select DataTable.Value("DAW","Global")

'Enter Rx Written Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Set DataTable.Value("Rx_Written_Date","Global") @@ hightlight id_;_989576_;_script infofile_;_ZIP::ssf4.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Type micTab

'Enter Prescribed Drug

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Drug").Set DataTable.Value("NDC","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find").Click
End If

If  JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").GetROProperty("enabled") Then
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").Click	
End If

'Enter Prescribed Quantity

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Qty.").Set DataTable.Value("Prescribed_Qty","Global") @@ hightlight id_;_31330009_;_script infofile_;_ZIP::ssf10.xml_;_

'Enter SIG Code

JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Set DataTable.Value("SIG_Code","Global") @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf8.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Type micTab

'Seacrh and select Route Of Administration

JavaWindow("Enterprise Pharmacy System").JavaEdit("Route of Administration").Set DataTable.Value("RouteOfAdmin_Code","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_4").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_4").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("ROA Search").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("ROA Search").JavaButton("Select").Click
End If

'Search and select prescriber @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf11.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescriber Last Name").Set DataTable.Value("PRESCRIBER_LAST_NAME","Global") @@ hightlight id_;_5189556_;_script infofile_;_ZIP::ssf12.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").Set DataTable.Value("PRESCRIBER_FIRST_NAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").Click
End If

'Enter Lot number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Lot Number").Set DataTable.Value("LOT_NUMBER","Global")

'Enter Drug Expiration Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Drug Expiration").Set DataTable.Value("DrugExpiration_Date","Global")
JavaWindow("Enterprise Pharmacy System").JavaEdit("Drug Expiration").Type micTab

'Enter Notes for this fill

JavaWindow("Enterprise Pharmacy System").JavaEdit("Notes For This Fill").Set DataTable.Value("ThisFillNotes","Global")
JavaWindow("Enterprise Pharmacy System").JavaEdit("Notes For This Fill").Type micTab

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click OK on Drug Validation Pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").JavaButton("OK").Click
End If

Wait(2)

'Generate Debug Images

JavaWindow("Virtual Scanner 2.0 -").JavaButton("Generate Debug Images").Click @@ hightlight id_;_30831015_;_script infofile_;_ZIP::ssf18.xml_;_

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient Last Name").Set DataTable.Value("PT_LASTNAME","Global") @@ hightlight id_;_18387578_;_script infofile_;_ZIP::ssf19.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient First Name").Set DataTable.Value("PT_FIRSTNAME","Global") @@ hightlight id_;_8115685_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

Wait(1)
		
JavaWindow("Enterprise Pharmacy System").JavaList("Site of Administration").Select DataTable.Value("Site_Of_Adminstration","Global")
	
Wait(1)
	
JavaWindow("Enterprise Pharmacy System").JavaList("Relationship Code").Select DataTable.Value("RelationshipCode","Global")

Wait(3)

datalastname = CStr(DataTable.Value("PT_LASTNAME","Global"))
datafirstname = CStr(DataTable.Value("PT_FIRSTNAME","Global"))
datamiddlename = CStr(DataTable.Value("PT_MIDDLENAME","Global"))

Wait(2)

getlastname = CStr(JavaWindow("Enterprise Pharmacy System").JavaEdit("Last Name").GetROProperty("value"))
getfirstname = CStr(JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").GetROProperty("value"))
getmiddlename = CStr(JavaWindow("Enterprise Pharmacy System").JavaEdit("Middle Name").GetROProperty("value"))

wait(2)

If datalastname = getlastname and datafirstname = getfirstname and datamiddlename = getmiddlename Then
	
	Reporter.ReportEvent micPass,"Patient Name","Patient Name gets auto-populated"
	Else
	Reporter.ReportEvent micFail,"Patient Name","Patient Name did not get auto-populated"
	
End If

'Check that transmit button is disabled

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micFail,"Transmit Button","Enabled"
	Else
	Reporter.ReportEvent micPass,"Transmit Button","Disabled - Birth Order is required"
End If

'Check that next button is disabled

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micFail,"Next Button","Enabled"
	Else
	Reporter.ReportEvent micPass,"Next Button","Disabled - Birth Order is required"
End If

'Populate Patient Birth Order

'Click on Patient File

If JavaWindow("Enterprise Pharmacy System").JavaButton("Patient File").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Patient File").Click
End If @@ hightlight id_;_20735409_;_script infofile_;_ZIP::ssf97.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaList("Multiple Birth").WaitProperty "visible",1

Wait(2)

'Select value from Birth Order dropdown

JavaWindow("Enterprise Pharmacy System").JavaList("Multiple Birth").Select DataTable.Value("Patient_MultipleBirth","Global")

Wait(1)

JavaWindow("Enterprise Pharmacy System").JavaList("Birth Order").Select DataTable.Value("Patient_BirthOrder","Global") @@ hightlight id_;_11295373_;_script infofile_;_ZIP::ssf107.xml_;_

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

Wait(5)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Address Not Found").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Address Not Found").JavaButton("OK").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Task").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Task").Click
End If

'Checkpoint for Birth Order on Patient Selection Screen

screenbirth = JavaWindow("Enterprise Pharmacy System").JavaEdit("Birth Order").GetROProperty("value")
strbirth = CStr(screenbirth)

If strbirth = DataTable.Value("Patient_BirthOrder","Global") Then
	
	Reporter.ReportEvent micPass,"Birth Order - Patient Selection Screen","The value has been auto-populated"
	Else
	Reporter.ReportEvent micFail,"Birth Order - Patient Selection Screen","The value is not correct"
	
End If

'Click on Edit Billing

If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Click
End If
 @@ hightlight id_;_17263703_;_script infofile_;_ZIP::ssf82.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Primary (1) Third Party").Select DataTable.Value("Billing_Option","Global") @@ hightlight id_;_17263703_;_script infofile_;_ZIP::ssf83.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Secondary (2) Third Party").Select DataTable.Value("Billing_Option2","Global") @@ hightlight id_;_29866389_;_script infofile_;_ZIP::ssf85.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Tertiary (3) Third Party").Select "#0" @@ hightlight id_;_29866389_;_script infofile_;_ZIP::ssf86.xml_;_

'Click on Save

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaButton("Save").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaButton("Save").Click
Else
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").Close
End If

'Click on Transmit

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Click
End If

'Checkpoint for Next button
	
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").WaitProperty "enabled",1
	
If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").GetROProperty("enabled") = 1 Then
	Reporter.ReportEvent micPass,"Next Button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Next Button","Disabled"
End If	
	
'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If



'Checkpoints for Batch Review Screen

JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").WaitProperty "visible",1

Wait(2)
	
	datalotnumber = CStr(DataTable.Value("LOT_NUMBER","Global"))
	datalastname = CStr(DataTable.Value("PT_LASTNAME","Global"))
	datafirstname = CStr(DataTable.Value("PT_FIRSTNAME","Global"))
	datasiteofadmin = CStr(DataTable.Value("Site_Of_Adminstration","Global"))
	datatpcarrier = CStr(DataTable.Value("ThirdParty1","Global"))

Wait(2)

	onscreenlotnumber = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Lot #"))
	onscreenpatlastname = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Patient Last Name"))
	onscreenpatfirstname = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Patient First Name"))
	onscreensiteofadmin = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Site of Administration"))
	onscreentpcarrier = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Carrier ID"))
	
	If datalotnumber = onscreenlotnumber and datalastname = onscreenpatlastname and datafirstname = onscreenpatfirstname and datasiteofadmin = onscreensiteofadmin and datatpcarrier = onscreentpcarrier Then
		
		Reporter.ReportEvent micPass,"Batch Review Screen","Details are correct"
		Else
		Reporter.ReportEvent micFail,"Batch Review Screen","Details are incorrect"
		
	End If
	
	
JavaWindow("Enterprise Pharmacy System").JavaButton("Back").WaitProperty "visible",1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If
'
'If WF_Value = "SDBP" Then
'	
'Dim strfirstname,strlastname
'
'strfirstname = DataTable.Value("PT_FIRSTNAME","Global")
'strlastname = DataTable.Value("PT_LASTNAME","Global")	
'
''Query to fetch the highest Rx Number  from table
'SQLquery = "Select max(EPS2.RX_SUMMARY.RX_NUMBER) RXNUM from EPS2.RX_SUMMARY,EPS2.PATIENT where EPS2.RX_SUMMARY.PATIENT_ID = EPS2.PATIENT.ID and EPS2.PATIENT.LAST_NAME = '"+ strlastname +"' and EPS2.PATIENT.FIRST_NAME = '"+ strfirstname +"'"
'
''Execute Query
'Set rcEPSRecordSet =  EPS2DBObject.Execute(SQLquery)
'
''assigning value to rxnumber
'rxnumber = rcEPSRecordSet.Fields("RXNUM")
'strrxnumber = CStr(rxnumber)
'
''Query to fetch rx/tx values from DB
'RxTxRelation = "select (EPS2.RX_TX.CONSENT_BY_RELATION_CD) RELATION,(EPS2.RX_TX.CONSENT_BY_MIDDLE_NAME) MNAME,(EPS2.RX_TX.CONSENT_BY_LAST_NAME) LNAME,(EPS2.RX_TX.CONSENT_BY_FIRST_NAME) FNAME from EPS2.RX_TX,EPS2.RX_SUMMARY where EPS2.RX_TX.RX_SUMMARY_ID = EPS2.RX_SUMMARY.ID and EPS2.RX_SUMMARY.RX_NUMBER = '"+ strrxnumber +"'"
'
''Execute Query
'Set rcEPSRecordSet =  EPS2DBObject.Execute(RxTxRelation)
'
''assigning value to rxtx fields
'
'rxtxrelation = rcEPSRecordSet.Fields("RELATION")
'rxtxmname = rcEPSRecordSet.Fields("MNAME")
'rxtxlname = rcEPSRecordSet.Fields("LNAME")
'rxtxfname = rcEPSRecordSet.Fields("FNAME")
'
'datarelation = CStr(DataTable.Value("CONSENT_GRANTED_BY_RELATION","Global"))
'
'If datarelation = rxtxrelation and datafirstname = rxtxfname and datalastname = rxtxlname and datamiddlename = rxtxmname Then
'	
'	Reporter.ReportEvent micPass,"DB Validation - Consent By Relation","All values are correct"
'	Else
'	Reporter.ReportEvent micFail,"DB Validation - Consent By Relation","Incorrect"
'	
'End If
'
'End If

'If WF_Value = "SDBP" Then
'	
''Validation for C131172
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Next Task").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Next Task").Click
'End If
'
''Enter User Authentication
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global") @@ hightlight id_;_32316305_;_script infofile_;_ZIP::ssf46.xml_;_
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")
'
'If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
'End If
'
'If JavaDialog("No tasks").Exist(15) Then
'
'Reporter.ReportEvent micPass,"Get Next - SDBP Rx(s)","User is not able to get Single Drug Batch Processing Rx(s) from Get Next"
'
'JavaDialog("No tasks").JavaButton("OK").Click
'End If @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf49.xml_;_
'
''Validation for C131175
'
''Navigate to Filecabinet>Patient>Rx/Tx
'
'JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Patient").JavaMenu("Rx/Tx").Select @@ hightlight id_;_167751_;_script infofile_;_ZIP::ssf51.xml_;_
'
''Search and select patient
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").Set DataTable.Value("PT_LASTNAME","Global")
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("First Name").Set DataTable.Value("PT_FIRSTNAME","Global") @@ hightlight id_;_8405362_;_script infofile_;_ZIP::ssf53.xml_;_
'
'If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Click
'End If
'
'Wait(3)
'
'If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
'End If
'
'
'
''Set Rx number
'
'JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx or Tx Number").WaitProperty "visible",1
'
'JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx or Tx Number").Set rxnumber @@ hightlight id_;_11558554_;_script infofile_;_ZIP::ssf56.xml_;_
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
'End If
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Tx Detail").Exist(10) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Tx Detail").Click
'End If
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Task Tracking").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Task Tracking").Click
'End If
'
''Check if DUR has run for Batch processed Rx
'
'getrows = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetROProperty("rows")
'	
'For i = 0 To getrows-1 Step 1
'	
'	gettaskdescription = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData(i,"Task Description")
'	getaction = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData(i,"Action")
'	
'If gettaskdescription = "Drug Utilization Review" and getaction = "complete" Then
'		Reporter.ReportEvent micPass,"SDBP - DUR","DUR has run for Batch Processed Rx"
'	Exit For
'End If
'
'Next
'
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Close Tx Detail").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Close Tx Detail").Click
'End If
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
'End If
'
'End If
'
'If WF_Value = "C131166" Then
'	
'Dim srfirstname,srlastname
'
'srfirstname = DataTable.Value("PT_FIRSTNAME","Global")
'srlastname = DataTable.Value("PT_LASTNAME","Global")	
'
''Query to fetch the highest Rx Number  from table
'SQLquery = "Select max(EPS2.RX_SUMMARY.RX_NUMBER) RXNUM from EPS2.RX_SUMMARY,EPS2.PATIENT where EPS2.RX_SUMMARY.PATIENT_ID = EPS2.PATIENT.ID and EPS2.PATIENT.LAST_NAME = '"+ srlastname +"' and EPS2.PATIENT.FIRST_NAME = '"+ srfirstname +"'"
'
''Execute Query
'Set rcEPSRecordSet =  EPS2DBObject.Execute(SQLquery)
'
''assigning value to rxnumber
'rxnumber = rcEPSRecordSet.Fields("RXNUM")
'strrxnumber = CStr(rxnumber)
'Wait(3)
'
'
''Navigate to Filecabinet>Patient>Rx/Tx
'
'JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Patient").JavaMenu("Rx/Tx").Select @@ hightlight id_;_22961763_;_script infofile_;_ZIP::ssf72.xml_;_
'
''Search and select patient
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").WaitProperty "visible",1
'
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").Set DataTable.Value("PT_LASTNAME","Global") @@ hightlight id_;_27700269_;_script infofile_;_ZIP::ssf73.xml_;_
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("First Name").Set DataTable.Value("PT_FIRSTNAME","Global") @@ hightlight id_;_21074664_;_script infofile_;_ZIP::ssf74.xml_;_
'
'If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Click
'End If
'
'JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").WaitProperty "enabled",1
' @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf76.xml_;_
' If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled")=1 Then
' 	JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
' End If
'
''Filter Rx by Batch Filled Rx option from dropdown
'
'JavaWindow("Enterprise Pharmacy System").JavaList("Filter by").Select "Batch Filled Rx" @@ hightlight id_;_20518300_;_script infofile_;_ZIP::ssf78.xml_;_
'
'JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx or Tx Number").Set strrxnumber @@ hightlight id_;_20109135_;_script infofile_;_ZIP::ssf79.xml_;_
'
''Click on Filter
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
'End If
'
''Validation for batch filled Rx on EPS
'
'rowvalue = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetROProperty("rows")
'
'If rowvalue > 0 Then
'	Reporter.ReportEvent micPass,"Filter By - Batch Filled Rx","User is able to filter by batch filled Rx"
'	Else
'	Reporter.ReportEvent micFail,"Filter By - Batch Filled Rx","User is unable to filter by batch filled Rx"
'End If
'
''Validation for rx status
'
'rxstatus = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Status")
'
'If rxstatus = "Active" Then
'	Reporter.ReportEvent micPass,"Rx Status","Active"
'	Else
'	Reporter.ReportEvent micFail,"Rx Status",""+ rxstatus +""
'End If
'
''Validation for tx status
'
'txstatus = JavaWindow("Enterprise Pharmacy System").JavaTable("Transaction Profile").GetCellData("#0","Status")
'
'If txstatus = "TP EX" Then
'	Reporter.ReportEvent micPass,"Tx Status","TP Exception"
'	Else
'	Reporter.ReportEvent micFail,"Tx Status",""+ txstatus +""
'End If
'
''DB Validation for batch filled rx
'
''Query to fetch the highest Rx Number  from table
'batchfilled = "select (EPS2.LINE_ITEM.LINE_ITEM_TYPE) ITEMTYPE from EPS2.LINE_ITEM,EPS2.RX_SUMMARY,EPS2.RX_TX where EPS2.LINE_ITEM.ID = EPS2.RX_TX.ID and EPS2.RX_TX.RX_SUMMARY_ID = EPS2.RX_SUMMARY.ID and EPS2.RX_SUMMARY.RX_NUMBER = '"+ strrxnumber +"'"
'
''Execute Query
'Set rcEPSRecordSet =  EPS2DBObject.Execute(batchfilled)
'
''assigning value to line item type
'itemtype = rcEPSRecordSet.Fields("ITEMTYPE")
'
'If itemtype = "18" Then
'	Reporter.ReportEvent micPass,"DB Validation - Line Item Type","18 - Batch Filled Rx"
'	Else
'	Reporter.ReportEvent micFail,"DB Validation - Line Item Type",""+ itemtype +""
'End If
'
'
''Click on Back to Home
'
'If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
'JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
'End If
'
'
'End If

Reporter.ReportEvent micDone,"Create New Batch","New Batch for single drug batch processing has been created" @@ hightlight id_;_26412249_;_script infofile_;_ZIP::ssf70.xml_;_
