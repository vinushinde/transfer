'..................................................................................................................................................
'Test Name : TP_Transmit_Queue_Complete
'Test Description : This test will transmit claim from TP Transmit Queue and complete it.
'..................................................................................................................................................

'Regression testing connection:
Set WshShell = CreateObject("WScript.Shell")
	Set WshSysEnv = WshShell.Environment("User")
	Set EPS2DBObject = CreateObject("ADODB.Connection") 
	Dim    vSchema 	, vEnvironment, vDSN
	vSchema  =  WshSysEnv("epsvSchema")
	vEnvironment = WshSysEnv("epsEnvironment")
	vDSN               =  WshSysEnv("vDSN")	
	vNhinID = WshSysEnv("vNHINID")
	WF_Value = WshSysEnv("WF")
	
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

Wait(3)

'Navigate to Queues>TP Transmit Queue

JavaWindow("Enterprise Pharmacy System").JavaMenu("Queues").JavaMenu("TP Transmit Queue (CTRL+P)").Select @@ hightlight id_;_9150902_;_script infofile_;_ZIP::ssf1.xml_;_

Wait(3)

'Enter Rx number
JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Number").Set rxnumber @@ hightlight id_;_23168202_;_script infofile_;_ZIP::ssf2.xml_;_

Wait(3)

'Click on Refresh

If JavaWindow("Enterprise Pharmacy System").JavaButton("Refresh").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Refresh").Click
End If

Wait(3)

If WF_Value = "C131177" Then
	
rowvalue = JavaWindow("Enterprise Pharmacy System").JavaTable("Third Party Claims").GetROProperty("rows")
	
	For i = 0 To rowvalue-1 Step 1
		
		getseq = JavaWindow("Enterprise Pharmacy System").JavaTable("Third Party Claims").GetCellData(i,"Seq.")
		
		If getseq = "Primary" Then
			JavaWindow("Enterprise Pharmacy System").JavaButton("View Transmit Detail").Click
		Exit For
		End If
		
	Next

Wait(2)

'Click on Page 7

If 	JavaWindow("Enterprise Pharmacy System").JavaButton("Page 7").Exist(15) Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Page 7").Click
End If
	
'Validation for C131177

dataincecntive = CStr(DataTable.Value("DEFAULT_INCENTIVE_FEE","Global"))
getprimaryincentive = CStr(JavaWindow("Enterprise Pharmacy System").JavaEdit("Incentive").GetROProperty("value"))

If dataincecntive = getprimaryincentive Then
	Reporter.ReportEvent micPass,"Default Incentive fee - Primary TP","Correct"
	Else
	Reporter.ReportEvent micFail,"Default Incentive fee - Primary TP","Incorrect"
End If

'Close transmit detail

If JavaWindow("Enterprise Pharmacy System").JavaButton("Close Transmit Detail").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Close Transmit Detail").Click
End If

'Validation for secondary TP

rowvaluesec = JavaWindow("Enterprise Pharmacy System").JavaTable("Third Party Claims").GetROProperty("rows")
	
	For j = 0 To rowvaluesec-1 Step 1
		
		getseqsecondary = JavaWindow("Enterprise Pharmacy System").JavaTable("Third Party Claims").GetCellData(j,"Seq.")
		
		If getseqsecondary = "Secondary" Then
			JavaWindow("Enterprise Pharmacy System").JavaTable("Third Party Claims").SelectRow j
			JavaWindow("Enterprise Pharmacy System").JavaButton("View Transmit Detail").Click
		Exit For
		End If
		
	Next

Wait(2)

'Click on Page 7

If 	JavaWindow("Enterprise Pharmacy System").JavaButton("Page 7").Exist(15) Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Page 7").Click
End If
	
'Validation for secondary TP for default incentive fee

getsecfee = JavaWindow("Enterprise Pharmacy System").JavaEdit("Incentive").GetROProperty("value")

If getsecfee = "---" Then
	Reporter.ReportEvent micPass,"Default Incentive Fee - Secondary TP","Correct"
	Else
	Reporter.ReportEvent micFail,"Default Incentive Fee - Secondary TP","Incorrect"
End If

'Close transmit detail

If JavaWindow("Enterprise Pharmacy System").JavaButton("Close Transmit Detail").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Close Transmit Detail").Click
End If

End If

'Enter Rx number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Number").Set "" @@ hightlight id_;_23168202_;_script infofile_;_ZIP::ssf2.xml_;_

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If


Reporter.ReportEvent micDone,"TP Transmit Queue","Claim has been completed" @@ hightlight id_;_20191822_;_script infofile_;_ZIP::ssf74.xml_;_
