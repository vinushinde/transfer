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

' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,"rt-qa-srv1-01","eps2app","prt7%r51ow"


WF_Value = WshSysEnv ("WF")

'Navigate to Tools>Utilities>Single Drug Batch Processing>Pharmacist Verification

JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Pharmacist Verification").Select @@ hightlight id_;_9402003_;_script infofile_;_ZIP::ssf1.xml_;_

If WF_Value = "C131166" Then

JavaWindow("Enterprise Pharmacy System").JavaTable("Results List").SelectRow "#0"

'Checkpoint for Top Table on Pharmacist Verification Screen

JavaWindow("Enterprise Pharmacy System").JavaObject("JTableHeader").Check CheckPoint("JTableHeader")

End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Reporter.ReportEvent micDone,"Pharmacist Verification","Work has been completed on Pharmacist Verification Screen"
