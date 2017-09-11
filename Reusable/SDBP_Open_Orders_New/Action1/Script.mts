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
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword


WF_Value = WshSysEnv ("WF")

'Navigate to Tools>Utilities>Single Drug Batch Processing

JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Open Orders").Select @@ hightlight id_;_3281681_;_script infofile_;_ZIP::ssf1.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaList("State").WaitProperty "visible",1

'Select workflow state of Rx

JavaWindow("Enterprise Pharmacy System").JavaCheckBox("Pending Batch Review Only_2").Set "ON"


'Click on Filter

If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
End If

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


If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Wait(2)



Reporter.ReportEvent micPass,"Open Order Screen","Pending Batch has been completed"
