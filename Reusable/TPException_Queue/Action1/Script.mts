'................................................................................................................................................. @@ hightlight id_;_23014116_;_script infofile_;_ZIP::ssf9.xml_;_
'Test Name : TPException_Queue_Retransmit
'Test Description : This test will perform the following functions:
                  '1. Navigate to TP Exception Queue(Queues>TP Exception Queue(Ctrl+T)).
                  '2. Search and select Patient.
                  '3. Click on Retransmit.
                  '4. Click on Complete
                  '5. Enter User Authentication
                  '6. Click on Back to Home
'................................................................................................................................................. @@ hightlight id_;_23014116_;_script infofile_;_ZIP::ssf9.xml_;_

''Regression testing connection:
Set WshShell = CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("User")
	Set EPS2DBObject = CreateObject("ADODB.Connection") 
	Dim    vSchema 	, vEnvironment, vDSN
	vSchema  =  WshSysEnv("epsvSchema")
	vEnvironment = WshSysEnv("epsEnvironment")
	vDSN               =  WshSysEnv("vDSN")	
	vNhinID = WshSysEnv("vNHINID")
	WF_Value = WshSysEnv("WF")
	Dim j
	j = WshSysEnv("rownum")


' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,"rt-qa-srv1-01","eps2app","prt7%r51ow"

'Import data from sheet into variables

Dim strfirstname,strlastname

strfirstname = DataTable.Value("PT_FIRSTNAME","Global")
strlastname = DataTable.Value("PT_LASTNAME","Global")	

'Query to fetch the highest Rx Number  from table
SQLquery = "Select max(EPS2.RX_SUMMARY.RX_NUMBER) RXNUM from EPS2.RX_SUMMARY,EPS2.PATIENT where EPS2.RX_SUMMARY.PATIENT_ID = EPS2.PATIENT.ID and EPS2.PATIENT.LAST_NAME = '"+ strlastname +"' and EPS2.PATIENT.FIRST_NAME = '"+ strfirstname +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(SQLquery)

'assigning value to rxnumber
rxnumber = rcEPSRecordSet.Fields("RXNUM")
strrxnumber = CStr(rxnumber)
Wait(3)


'Navigate to Queues>TP Exception Queue(CTRL+T)

JavaWindow("Enterprise Pharmacy System").JavaMenu("Queues").JavaMenu("TP Exception Queue (CTRL+T)").Select @@ hightlight id_;_1777075_;_script infofile_;_ZIP::ssf10.xml_;_
 @@ hightlight id_;_8353520_;_script infofile_;_ZIP::ssf12.xml_;_
 Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Number").Set rxnumber

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Number").Type micReturn

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(5) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

If WF_Value = "C131166" Then

'Click on Edit Billing button	
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf24.xml_;_
 If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Exist(10) Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Click 
 End If

'Select TP that will pay

JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Primary (1) Third Party").Select DataTable.Value("Billing_Option2","Global") @@ hightlight id_;_22639975_;_script infofile_;_ZIP::ssf26.xml_;_

'Click on Save

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaButton("Save").Click
End If

'Click on No on other payer reject codes

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Other Payer Reject Codes").Exist Then
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Other Payer Reject Codes").JavaButton("No").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Message").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Message").JavaButton("OK").Click
End If

End If


'Click on Retransmit

If JavaWindow("Enterprise Pharmacy System").JavaButton("Retransmit").Exist(5) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Retransmit").Click	
End If



'Click on Complete @@ hightlight id_;_1508064_;_script infofile_;_ZIP::ssf16.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Exist(5) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Click
End If

'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global") @@ hightlight id_;_32191517_;_script infofile_;_ZIP::ssf18.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(5) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Printing Error Dialog

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Printing Error").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Printing Error").JavaButton("OK").Click
End If

'Blank Rx Number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Number").Set ""

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(5) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click	
End If

If WF_Value = "C131166" Then
	
'Navigate to Filecabinet>Patient>Rx/Tx

JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Patient").JavaMenu("Rx/Tx").Select @@ hightlight id_;_15458889_;_script infofile_;_ZIP::ssf61.xml_;_

'Search and select Patient

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").Set DataTable.Value("PT_LASTNAME","Global") @@ hightlight id_;_12288169_;_script infofile_;_ZIP::ssf62.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("First Name").Set DataTable.Value("PT_FIRSTNAME","Global") @@ hightlight id_;_3632564_;_script infofile_;_ZIP::ssf63.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Exist(15) Then

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

End If

'Enter rx number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx or Tx Number").Set rxnumber @@ hightlight id_;_5258319_;_script infofile_;_ZIP::ssf66.xml_;_

Wait(3)

'Click on Filter

If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
End If

Wait(3)

'Validation for status of Tx

txstatus = JavaWindow("Enterprise Pharmacy System").JavaTable("Transaction Profile").GetCellData("#0","Status")

If txstatus = "DV" Then
	Reporter.ReportEvent micPass,"Tx Status","Tx is in Data Verification"
	Else
	Reporter.ReportEvent micFail,"Tx Status","Tx is in "+ txstatus +""
End If
 @@ hightlight id_;_11349887_;_script infofile_;_ZIP::ssf71.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click @@ hightlight id_;_8561749_;_script infofile_;_ZIP::ssf72.xml_;_
	
End If


Reporter.ReportEvent micDone,"TP Exception Retransmit","The claim has been retransmitted successfully" @@ hightlight id_;_13427946_;_script infofile_;_ZIP::ssf57.xml_;_
