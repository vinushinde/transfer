'..............................................................................................................................................

'Test Name : RevertChanges_ECC

'Test Description : This test will revert all the changes made on ECC

'..............................................................................................................................................

'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbuser = WshSysEnv("vDbuser")
Dbpwd = WshSysEnv("vDbpwd")

vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

'Call actions

LaunchIE(vurl)
Wait(Iteration_Wait)
If Browser("Enterprise Control Center_2").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Certificate Error: Navigation").Link("Continue to this website").Click
End If
Wait(Iteration_Wait)

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center_3").WebEdit("username").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center_3").WebEdit("username").Set username
End If

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center_3").WebEdit("password").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center_3").WebEdit("password").Set password
End If

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center_3").WebArea("OK").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center_3").WebArea("OK").Click
End If

Wait(2)

'Navigate to Administration>User Security>Role Settings>EPS Role Settings

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id105:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id105:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Search for Inventory Roles

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set "Inv" @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch")_;_script infofile_;_ZIP::ssf2.xml_;_

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click
End If


'Add all the inventory roles

Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").WaitProperty "visible",1

Set oDesc = Description.Create()
oDesc("micclass").Value = "Link"
Set Links = Browser("Enterprise Control Center").Page("Enterprise Control Center").ChildObjects(oDesc)
TotalLinks = Links.count
print "TotalLinks:"&TotalLinks

For Iterator =  TotalLinks To 1 Step -1
print "Iterator:"&Iterator
Browser("Enterprise Control Center_4").Page("Enterprise Control Center").Link("Remove").Click
Wait(1)
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set "Inv"
Wait(1)
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click
Wait(1)
Next

'Click on Home button

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebTable("Home").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebTable("Home").Click
End If

'Navigate to Administration>User Security>Role Settings>EPS Role Settings

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id105:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id105:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Search for Will Call Role

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set "Will Call" @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch")_;_script infofile_;_ZIP::ssf2.xml_;_

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click
End If

Wait(2)

If Browser("Enterprise Control Center_4").Page("Enterprise Control Center").Link("Remove").Exist(15) Then
Browser("Enterprise Control Center_4").Page("Enterprise Control Center").Link("Remove").Click
End If


'Click on Home button

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebTable("Home").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebTable("Home").Click
End If

'Close Internet explorer

Browser("Enterprise Control Center").Close


'Ping ECC

PingECC()

Reporter.ReportEvent micPass,"Revert Changes","All changes have been reverted from ECC"
