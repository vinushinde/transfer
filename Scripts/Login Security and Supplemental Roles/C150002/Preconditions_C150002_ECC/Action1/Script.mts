'......................................................................................................................................

'Test Name : Preconditions_C150002_ECC

'Test Description : This test will setup all the preconditions on ECC

'......................................................................................................................................


'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

vurl="https://"&Serverip&":58442/ecc/login.jsp"


'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

'Call actions 


LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password

Wait(3)

'Add 3 supplemental roles on ECC

'Navigate to Administration>User Security>Role Settings>EPS Role Settings

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id105:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id105:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Add first supplemental role

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").WaitProperty "visible",1

Wait(2)

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set DataTable.Value("SupplementalRole1","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch")_;_script infofile_;_ZIP::ssf2.xml_;_

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click
End If

Wait(2)

If Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Click
End If

Wait(2)

'Add second supplemental role

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set DataTable.Value("SupplementalRole2","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch")_;_script infofile_;_ZIP::ssf5.xml_;_

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click
End If

Wait(2)

If Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Click
End If

Wait(2)

'Add third supplemental role

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set DataTable.Value("SupplementalRole3","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch")_;_script infofile_;_ZIP::ssf8.xml_;_

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click
End If

Wait(2)

If Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Click
End If

'Navigate to Home 

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("WebTable").Click
End If
 @@ hightlight id_;_32698451_;_script infofile_;_ZIP::ssf43.xml_;_
If Browser("Certificate Error: Navigation").Exist(15) Then
Browser("Certificate Error: Navigation").Close
End If



 Reporter.ReportEvent micDone,"Preconditions","All preconditions on ECC have been setup" @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Control Center").WebArea("https://192.168.109.121:58442/ecc/i")_;_script infofile_;_ZIP::ssf47.xml_;_
