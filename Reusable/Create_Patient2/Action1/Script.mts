'..................................................................................................................................

'Test Name : Create_Patient

'Test Description : This test will create a new patient if it is not already present on the client

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



'Navigate to Filecabinet>Patient>Information

JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Patient").JavaMenu("Information").Select @@ hightlight id_;_11148416_;_script infofile_;_ZIP::ssf1.xml_;_

'Search patient

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("Last Name").Set DataTable.Value("PT_LASTNAME2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaEdit("First Name").Set DataTable.Value("PT_FIRSTNAME2","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("find").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaStaticText("No Record Found(st)").Exist(20) Then

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Add New").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Add New").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("Address").WaitProperty "visible",1 @@ hightlight id_;_10651176_;_script infofile_;_ZIP::ssf8.xml_;_

If WF_Value = "SDBPSprint1" or WF_Value = "SDBP2608" Then
	
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("Middle Name").Set DataTable.Value("PT_MIDDLENAME","Global")
	
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("Address").Set DataTable.Value("PT_ADDRESS2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("City").Set DataTable.Value("PT_CITY2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaList("State").Select DataTable.Value("PT_STATE2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("ZIP Code").Set DataTable.Value("PT_ZIPCODE2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaCheckBox("Do not Verify This Address").Set "ON" @@ hightlight id_;_20577733_;_script infofile_;_ZIP::ssf14.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("Home Phone Number").Set DataTable.Value("PT_PHONE2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaEdit("Date of Birth").Set DataTable.Value("PT_DOB2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaList("Gender").Select DataTable.Value("PT_GENDER2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaRadioButton("No Known Drug Allergies").Set @@ hightlight id_;_10827672_;_script infofile_;_ZIP::ssf18.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaRadioButton("No Known Medical History").Set @@ hightlight id_;_12579663_;_script infofile_;_ZIP::ssf19.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaButton("Next").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaDialog("Add Patient").JavaButton("Next").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(20) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").Close

End If

Reporter.ReportEvent micDone,"Create New Patient","New patient has been created" @@ hightlight id_;_30512343_;_script infofile_;_ZIP::ssf21.xml_;_
