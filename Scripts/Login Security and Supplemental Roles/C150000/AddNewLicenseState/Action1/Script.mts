'...........................................................................................................................................

'Test Name : AddNewLicenseState

'...........................................................................................................................................


'Navigate to Administration>User Security>User Licenses>EPS User License Requirements

If Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebElement("header-form:j_id109:anchor").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebElement("header-form:j_id109:anchor").Click
End If

'Select State for which you need to edit the license

rowval = Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebTable("User License State Requirements").GetROProperty("rows")

For i = 0 To rowval Step 1
	
	celdt = Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebTable("User License State Requirements").GetCellData(i,"1")
	
	If Instr(celdt,"TEXAS")>0 Then
	
	Set tem = Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebTable("User License State Requirements").ChildItem(i,2,"Link",0)	
	tem.Click
	Exit For
	End If
Next

'Add custom license to RPh License

Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebList("userLicenseReqForm:licenseType").WaitProperty "visible",1


Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebList("userLicenseReqForm:licenseType").Select DataTable.Value("LicenseTypeName","Global")
Wait(2)
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebList("userLicenseReqForm:userTypeCombo").Select "RPh" @@ hightlight id_;_Browser("EPS - User License Requirement 2").Page("EPS - User License Requirement").WebList("userLicenseReqForm:userTypeCombo")_;_script infofile_;_ZIP::ssf37.xml_;_
Wait(2)

If Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Click
End If

'Add RPh license to RPh License

Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebList("userLicenseReqForm:licenseType").WaitProperty "visible",1


Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebList("userLicenseReqForm:licenseType").Select "RPh"
Wait(2)
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebList("userLicenseReqForm:userTypeCombo").Select "RPh" @@ hightlight id_;_Browser("EPS - User License Requirement 2").Page("EPS - User License Requirement").WebList("userLicenseReqForm:userTypeCombo")_;_script infofile_;_ZIP::ssf37.xml_;_
Wait(2)

If Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Click
End If

'Click on Home button

If Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebElement("WebTable").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebElement("WebTable").Click
End If



Reporter.ReportEvent micDone,"License on state requirements screen","Both licenses have been added on state requirement screen" @@ hightlight id_;_Browser("EPS - User License Requirement 2").Page("EPS - User License Requirement").WebElement("WebTable")_;_script infofile_;_ZIP::ssf35.xml_;_
