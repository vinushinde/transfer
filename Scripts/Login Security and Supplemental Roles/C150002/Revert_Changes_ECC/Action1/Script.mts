'.......................................................................................................................................

'Test name : Preconditions_C150002_ECC

'Test Description : This test case will set all the preconditions on EPS

'.......................................................................................................................................


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


'Call actions

'Expire 1st and 3rd role from ECC by removing from EPS role settings

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")


Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")


LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password

Wait(3)


'Navigate to Administration>User Security>Role Settings>EPS Role Settings

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebElement("header-form:j_id105:anchor").Click

'Remove first supplemental role

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebEdit("searchForm:roleSearch").WaitProperty "visible",1

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebEdit("searchForm:roleSearch").Set DataTable.Value("SupplementalRole1","Global")

If Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebButton("Search").Exist(15) Then
Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebButton("Search").Click
End If

Wait(2)

If Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").Link("Remove").Exist(15) Then
Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").Link("Remove").Click
End If

'Remove third supplemental role

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebEdit("searchForm:roleSearch").WaitProperty "visible",1

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebEdit("searchForm:roleSearch").Set DataTable.Value("SupplementalRole3","Global")

If Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebButton("Search").Exist(15) Then
Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebButton("Search").Click
End If

Wait(2)

If Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").Link("Remove").Exist(15) Then
Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").Link("Remove").Click
End If


'Navigate to Home 

If Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebElement("WebTable").Exist(15) Then
Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebElement("WebTable").Click
End If

If Browser("Certificate Error: Navigation").Exist(15) Then
Browser("Certificate Error: Navigation").Close
End If


Reporter.ReportEvent micDone,"Revert Changes - ECC","All changes have been reverted from ECC"
