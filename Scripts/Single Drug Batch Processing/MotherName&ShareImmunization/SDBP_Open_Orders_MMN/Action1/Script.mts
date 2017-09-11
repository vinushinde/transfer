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
 @@ hightlight id_;_8653509_;_script infofile_;_ZIP::ssf31.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaList("State").Select DataTable.Value("Workflow_State","Global") @@ hightlight id_;_8653509_;_script infofile_;_ZIP::ssf32.xml_;_


'Click on Filter

If JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Filter").Click
End If


'Cancel all Rxs

resultrows = JavaWindow("Enterprise Pharmacy System").JavaTable("Results List_2").GetROProperty("rows")

For i = resultrows-1 To 0 Step -1
	
	JavaWindow("Enterprise Pharmacy System").JavaTable("Results List_2").SelectRow i
	
	batchrows = JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").GetROProperty("rows")
	
	For j = batchrows-1 To 0 Step -1
		JavaWindow("Enterprise Pharmacy System").JavaTable("Batch Details").SelectRow j
		
		If JavaWindow("Enterprise Pharmacy System").JavaButton("Cancel Rx").Exist(15) Then
		JavaWindow("Enterprise Pharmacy System").JavaButton("Cancel Rx").Click
		End If
		
		If JavaWindow("Enterprise Pharmacy System").JavaDialog("Cancel Prescription").Exist(10) Then
			JavaWindow("Enterprise Pharmacy System").JavaDialog("Cancel Prescription").JavaButton("Yes").Click
		End If
		
		Wait(2)
		
	Next
	
Next


'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(10) Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If


Reporter.ReportEvent micPass,"Open Order Screen","All Rxs have been cancelled"
