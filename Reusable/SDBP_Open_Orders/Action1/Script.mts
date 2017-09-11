'...........................................................................................................................................

'Test Name : SDBP_Open_Orders

'Test Description : This test will navigate to Tools>Utilities>Single Drug Batch Processing>Open Orders

'Author : Kashish Ambwani

'Date Modified : 9 June 2017

'...........................................................................................................................................

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


WF_Value = WshSysEnv ("WF")

'Navigate to Tools>Utilities>Single Drug Batch Processing

JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Open Orders").Select @@ hightlight id_;_3281681_;_script infofile_;_ZIP::ssf1.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaList("State").WaitProperty "visible",1

'Select workflow state of Rx

If WF_Value = "C131179" Then

	JavaWindow("Enterprise Pharmacy System").JavaCheckBox("Pending Batch Review Only_2").Set "ON"
	ElseIf WF_mode = "WORKFLOW" Then
	JavaWindow("Enterprise Pharmacy System").JavaList("State").Select DataTable.Value("WORKFLOW_STATE","Global")
	ElseIf WF_mode = "RAPIDFILL" Then
	JavaWindow("Enterprise Pharmacy System").JavaList("State").Select DataTable.Value("STATE","Global")
End If

If WF_Value = "C131181" or WF_Value = "C131179" or WF_Value = "C131180" Then

JavaWindow("Enterprise Pharmacy System").JavaEdit("NDC").Set DataTable.Value("NDC","Global")

End If

'Click on Filter

If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
End If

If WF_Value = "C131181" Then

rowval = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetROProperty("rows")

For i = rowval-1 To 0 Step -1

JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").SelectRow i

If JavaWindow("Enterprise Pharmacy System").JavaButton("Cancel Rx").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Cancel Rx").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Cancel Prescription").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Cancel Prescription").JavaButton("Yes").Click
End If

Reporter.ReportEvent micPass,"Cancel Rx(s)","All Rx(s) have been cancelled"

Next

ElseIf WF_Value = "C131179" Then

rowvalue = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List_2").GetROProperty("rows")

For j = rowvalue-1 To 0 Step -1
	
JavaWindow("Enterprise Pharmacy System").JavaTable("Results List_2").SelectRow j

'Click on Complete

JavaWindow("Enterprise Pharmacy System").JavaButton("Complete_2").Click
	
'Enter User Authentication

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaEdit("User Name").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaEdit("Password").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaButton("OK").Click
End If

Reporter.ReportEvent micPass,"Rx - Status","Rx is in Pharmacist Verification"


Next

ElseIf WF_Value = "C131180" Then

Wait(3)

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

'Select row for COB

getrows = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetROProperty("rows")

For i = 0 To getrows-1 Step 1
	
	getrx = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData(i,"RX Number")
	
	If getrx = strrxnumber Then
		JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").SelectRow i
		Exit For
	End If
	
Next
	
'Click on Review COB @@ hightlight id_;_17571395_;_script infofile_;_ZIP::ssf22.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("Review CoB").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Review CoB").Click
End If

'Click on complete

If JavaWindow("Enterprise Pharmacy System").JavaButton("Complete_2").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Complete_2").Click
End If

'Enter User Authentication


JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global") @@ hightlight id_;_28856308_;_script infofile_;_ZIP::ssf25.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaButton("OK").Exist(15) Then
Reporter.ReportEvent micPass,"COB Complete","User is able to complete COB"
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System_2").JavaButton("OK").Click
Else
Reporter.ReportEvent micFail,"COB Complete","User is unable to compelete COB"
End If

End If
 @@ hightlight id_;_7196609_;_script infofile_;_ZIP::ssf13.xml_;_
If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If


Reporter.ReportEvent micDone,"SDBP-Open Order","Work on Open Orders Screen has been completed" @@ hightlight id_;_24959569_;_script infofile_;_ZIP::ssf21.xml_;_
