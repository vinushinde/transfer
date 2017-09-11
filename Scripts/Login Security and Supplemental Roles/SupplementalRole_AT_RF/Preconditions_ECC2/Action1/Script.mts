'..............................................................................................................................................

'Test Name : Preconditions_C150045

'Test Description : This test will perform all the preconditions for this test case

'Date Modified : 20 June 2017

'..............................................................................................................................................

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
Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Click
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

If Browser("Enterprise Control Center_3").Page("Enterprise Control Center").Link("Add").Exist(15) Then
Browser("Enterprise Control Center_3").Page("Enterprise Control Center").Link("Add").Click
End If


'Click on Home button

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebTable("Home").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebTable("Home").Click
End If

'Close Internet Explorer

If Browser("Enterprise Control Center").Exist(15) Then
Browser("Enterprise Control Center").Close
End If


Reporter.ReportEvent micPass,"Preconditions","All preconditions have been set"
