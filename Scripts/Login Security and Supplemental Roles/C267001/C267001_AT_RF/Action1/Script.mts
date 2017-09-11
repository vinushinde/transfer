'..................................................................................................................................

'Test Name : Verify alert message while logging into EPS when Pharmacist of Record Alert is set to Yes

'Test Description : Verify alert message while logging into EPS when Pharmacist of Record Alert is set to Yes

'TestRail ID : C267001

'Author : Kashish Ambwani

'Date Modified : 16 August 2017

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
	WF_mode = WshSysEnv("vMODE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

WshSysEnv ("WF") = "C267001"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Import variables from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

empid2 = RandomUserType(6)
password2 = DataTable.Value("Password2","Global")
userlastname2 = RandomUserType(4)
userfirstname2 = RandomUserType(4)
userinitials2 = RandomUserType(3)
usergroup2 = DataTable.Value("User_Group","Global")
usertype2 = DataTable.Value("UserUserType","Global")

'Call actions here

PharmacistRecordAlert_ApplicationSettings("Yes")
Wait(Iteration_Wait)
CreateNewUserEPS_New empid2,userlastname2,userfirstname2,userinitials2,usergroup2,empid2,password2,usertype2,username,password
Wait(Iteration_Wait)
AddLicenseUserEPS(empid2)
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin empid2,password2
Wait(Iteration_Wait)

'**********************************************************Start of validation 1********************************************************

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").Exist Then
	Reporter.ReportEvent micPass,"Change RPh Of Record Dialog","Displayed"
	Else
	Reporter.ReportEvent micFail,"Change RPh Of Record Dialog","Not-Displayed"
End If

'Validations for all 3 buttons

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").JavaButton("Assign to Me").GetROProperty("displayed")=1 Then
	Reporter.ReportEvent micPass,"Assign To Me button","Button is displayed"
	Else
	Reporter.ReportEvent micFail,"Assign To Me button","Button is not displayed"
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").JavaButton("Change Rph of Record").GetROProperty("displayed")=1 Then
	Reporter.ReportEvent micPass,"Change RPh Of Record button","Button is displayed"
	Else
	Reporter.ReportEvent micFail,"Change RPh Of Record button","Button is not displayed"
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").JavaButton("Do not Change").GetROProperty("displayed")=1 Then
	Reporter.ReportEvent micPass,"Do Not Change button","Button is displayed"
	Else
	Reporter.ReportEvent micFail,"Do Not Change button","Button is not displayed"
End If

'Validate if focus is on Do not change button

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").JavaButton("Do not Change").GetROProperty("focused")=1 Then
	Reporter.ReportEvent micPass,"Do Not change button - Focus","Focus is on Do not Change button"
	Else
	Reporter.ReportEvent micFail,"Do Not change button - Focus","Focus is not on Do not Change button"
End If

'Click on Do not Change Button

JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").Type micAlt+"D"

Reporter.ReportEvent micPass,"Do not change button hot key","Button has been pressed by using the respective hotkey"

Wait(3)

readlable = JavaWindow("Enterprise Pharmacy System").JavaStaticText("<html>Pharmacist of Record:").GetROProperty("label")
'msgbox readlable
start = Instr(readlable,"<b>")
endstr =Instr(readlable,"</b>")
startpos = start+3
mystring = Mid(readlable,startpos,endstr-startpos)

strstring = CStr(mystring)

Reporter.ReportEvent micPass,"Pharmacist Of Record after clicking on Do Not Change button"," Pharmacist Of Record is  "+ strstring +""

'**************************************************End of validation 1*******************************************************************

Wait(Iteration_Wait)

EPSLogout()
Wait(Iteration_Wait)
EPSLogin empid2,password2
Wait(Iteration_Wait)

'**************************************************Start of validation 2*****************************************************************

'Click on Assign to Me button

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").Exist(15) Then
	
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").Type micAlt+"A"
	
	Reporter.ReportEvent micPass,"Assign To Me button","Assign to me button has been pressed by the respective hot key"
	
End If

userfullname2 = userfirstname2&" "&userlastname2
strfullname2 = CStr(userfullname2)

readlable2 = JavaWindow("Enterprise Pharmacy System").JavaStaticText("<html>Pharmacist of Record:").GetROProperty("label")
'msgbox readlable
start2 = Instr(readlable2,"<b>")
endstr2 =Instr(readlable2,"</b>")
startpos2 = start2+3
mystring2 = Mid(readlable2,startpos2,endstr2-startpos2)

strstring2 = CStr(mystring2)

If strstring2 = strfullname2 Then
	Reporter.ReportEvent micPass,"Pharmacist Of Record after clicking on Assign To me","Correct;Pharmacist is "+ strstring2 +""
	Else
	Reporter.ReportEvent micFail,"Pharmacist Of Record after clicking on Assign To me","Incorrect;Expected : "+ strfullname2 +";Observed : "+ strstring2 +""
End If

'************************************************************End of validation 2***********************************************************

EPSLogout()
Wait(Iteration_Wait)
EPSLogin empid2,password2
Wait(Iteration_Wait)

'***************************************************Start of validation 3******************************************************************

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").Exist(15) Then
	
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Change Rph Of Record").Type micAlt+"C"
	
	Reporter.ReportEvent micPass,"Change RPh of Record button","Change RPh Of Record has been pressed using the respective hot key"
	
End If

Wait(2)

screenpharmacist = JavaWindow("Enterprise Pharmacy System").JavaEdit("Pharmacist of Record").GetROProperty("value")
strscreen = CStr(screenpharmacist)

If strscreen = strfullname2 Then
	Reporter.ReportEvent micPass,"Pharmacist Of record on Change Rph Of Record screen","Correct"
	Else
	Reporter.ReportEvent micFail,"Pharmacist Of record on Change Rph Of Record screen","Expected :"+ strfullname2 +";Obserevd : "+ strscreen +""
End If

'Select pharmacist from list

rowval = JavaWindow("Enterprise Pharmacy System").JavaTable("Available Pharmacists").GetROProperty("rows")

For i = 0 To rowval-1 Step 1
	
	name = JavaWindow("Enterprise Pharmacy System").JavaTable("Available Pharmacists").GetCellData(i,"Pharmacist")
	If name = "anil kumar" Then
		JavaWindow("Enterprise Pharmacy System").JavaTable("Available Pharmacists").SelectRow i
	Exit For
	
	End If
	
Next

'Click on Save

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

Wait(2)
 @@ hightlight id_;_20806396_;_script infofile_;_ZIP::ssf21.xml_;_
'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Wait(2)


JavaWindow("Enterprise Pharmacy System").JavaStaticText("<html>Pharmacist of Record:").Check CheckPoint("<html>Pharmacist of Record: <b>anil kumar</b></html>(st)") @@ hightlight id_;_1360840_;_script infofile_;_ZIP::ssf23.xml_;_

'*********************************************End of validation 3***************************************************************************


PharmacistRecordAlert_ApplicationSettings("No")
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username,password


Reporter.ReportEvent micDone,"Test Case  - C267001","Test case has been executed successfully"
