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

'Import data from sheet into variables

Dim strfirstname,strlastname

strfirstname = DataTable.Value("PT_FIRSTNAME","Global")
strlastname = DataTable.Value("PT_LASTNAME","Global")	

'Query to fetch the highest Rx Number  from table
SQLquery = "Select max("+ vSchema +"RX_SUMMARY.RX_NUMBER) RXNUM from "+ vSchema +"RX_SUMMARY,"+ vSchema +"PATIENT where "+ vSchema +"RX_SUMMARY.PATIENT_ID = "+ vSchema +"PATIENT.ID and "+ vSchema +"PATIENT.LAST_NAME = '"+ strlastname +"' and "+ vSchema +"PATIENT.FIRST_NAME = '"+ strfirstname +"'"

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



'Click on Edit Billing button	
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf24.xml_;_
 If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Exist(10) Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Click 
 End If

'Select TP that will pay


JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Primary (1) Third Party").Select DataTable.Value("Billing_Option2","Global")
Wait(1)
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Secondary (2) Third Party").Select DataTable.Value("Billing_Option3","Global")
Wait(1)

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


If WF_mode = "RAPIDFILL" Then
	
If JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If
	
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



Reporter.ReportEvent micDone,"TP Exception Retransmit","The claim has been retransmitted successfully" @@ hightlight id_;_13427946_;_script infofile_;_ZIP::ssf57.xml_;_
