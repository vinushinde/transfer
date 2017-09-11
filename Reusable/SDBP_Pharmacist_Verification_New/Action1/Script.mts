'..............................................................................................................................................

'Test Name : This test will navigate to the Tools>Utilities>Single Drug Batch Processing>Pharmacist Verification screen

'Author : Kashish Ambwani

'..............................................................................................................................................
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

'Navigate to Tools>Utilities>Single Drug Batch Processing>Pharmacist Verification

JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Pharmacist Verification").Select @@ hightlight id_;_9402003_;_script infofile_;_ZIP::ssf1.xml_;_

Wait(3)

JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").SelectRow "#0"


'Validations on Pharmacist Verification Screen

screenlotnumber = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Lot #"))
screenfilldate = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Fill Date"))
screentotal = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").GetCellData("#0","Total"))

Wait(2)

screenbatchlot = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData("#0","Lot #"))
screenbatchlastname = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData("#0","Patient Last Name"))
screenbatchfirstname = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData("#0","Patient First Name"))
screenbatchdob = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData("#0","Date of Birth"))
screenbatchcarrier = CStr(JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData("#0","Carrier ID"))

If screenlotnumber = screenbatchlot and screenbatchlot = DataTable.Value("LOT_NUMBER","Global") Then
	Reporter.ReportEvent micPass,"Lot Number - Pharmacist Verification Screen","Correct"
	Else
	Reporter.ReportEvent micFail,"Lot Number - Pharmacist Verification Screen","Incorrect"
End If

Reporter.ReportEvent micPass,"Fill Date - Pharmacist Verification",""+ screenfilldate +""

Reporter.ReportEvent micPass,"Fill Date - Pharmacist Verification",""+ screentotal +""

If screenbatchfirstname = CStr(DataTable.Value("PT_FIRSTNAME","Global")) and screenbatchlastname = CStr(DataTable.Value("PT_LASTNAME","Global")) Then
	Reporter.ReportEvent micPass,"Patient Details - Pharmacist Verification","Correct"
	Else
	Reporter.ReportEvent micFail,"Patient Details - Pharmacist Verification","Incorrect"
End If

Reporter.ReportEvent micPass,"Patient DOB - Pharmacist Verification Screen",""+ screenbatchdob +""

Wait(2)

'Check select all checkbox functionality

'Check Select All checkbox

JavaWindow("Enterprise Pharmacy System").JavaCheckBox("Select All").Set "ON" @@ hightlight id_;_28366602_;_script infofile_;_ZIP::ssf11.xml_;_

rowval = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetROProperty("rows")

For i = 0 To rowval-1 Step 1
	
	checkvalue = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData(i,"Print Receipt")
	
stri = CStr(i)
	
	If checkvalue = 1 Then
		Reporter.ReportEvent micPass,"Select All checkbox checked : Row number : "+ stri +" - Print Receipt checkbox","Checked"
		Else
		Reporter.ReportEvent micFail,"Select All checkbox checked : Row number : "+ stri +" - Print Receipt checkbox","Unchecked"
	End If
	
Next

Wait(2)

'Uncheck Select All checkbox

JavaWindow("Enterprise Pharmacy System").JavaCheckBox("Select All").Set "OFF" @@ hightlight id_;_28366602_;_script infofile_;_ZIP::ssf12.xml_;_

rowvalue = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetROProperty("rows")

For j = 0 To rowvalue-1 Step 1
	
	checkvalueoff = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData(j,"Print Receipt")
	
strj = CStr(j)
	
	If checkvalueoff = 0 Then
		Reporter.ReportEvent micPass,"Select All checkbox unchecked : Row number : "+ strj +" - Print Receipt checkbox","Unchecked"
		Else
		Reporter.ReportEvent micFail,"Select All checkbox unchecked : Row number : "+ strj +" - Print Receipt checkbox","Checked"
	End If
	
Next

Wait(2)

'Check print receipt checkbox for first row

JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").SetCellData "#0","Print Receipt","1" @@ hightlight id_;_23857770_;_script infofile_;_ZIP::ssf13.xml_;_

Wait(2)

'Check if checkbox for first row is checked

checkfirst = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetCellData("#0","Print Receipt")

If checkfirst = 1 Then
	Reporter.ReportEvent micPass,"Indiviual checkbox validation","Checked"
	Else
	Reporter.ReportEvent micFail,"Indiviual checkbox validation","Unchecked"
End If

Wait(2)

'Uncheck checkbox for first row

JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").SetCellData "#0","Print Receipt","0" @@ hightlight id_;_23857770_;_script infofile_;_ZIP::ssf13.xml_;_

Wait(2)

'Click on complete

If JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Click
End If

'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global") @@ hightlight id_;_24108815_;_script infofile_;_ZIP::ssf7.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Reporter.ReportEvent micDone,"Pharmacist Verification","Work has been completed on Pharmacist Verification Screen"
