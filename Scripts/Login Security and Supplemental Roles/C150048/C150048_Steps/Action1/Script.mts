'...........................................................................................................................................

'Test Name : C150048_Steps

'...........................................................................................................................................


'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

'Import variables from sheet

licensetypname = RandomUserType(5)
strlicensetypename = CStr(licensetypname)

'Add new license type

AddNewLicenseType_ECC(licensetypname)
Wait(Iteration_Wait)



'Navigate to Administration>User Security>User Licenses>EPS User License Requirements

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id109:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id109:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Click on Add New

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Add New").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Add New").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Add New").Click
End If

'Enter state name and license and user type

Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").Select "#1"

statename = Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").GetROProperty("value")
strstate = CStr(statename)

Browser("EPS - User License Requirement").Page("EPS - User License Requirement").WebList("userLicenseReqForm:licenseType").Select strlicensetypename

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebList("userLicenseReqForm:userTypeCombo").Select DataTable.Value("UserType","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS - User License Requirement 2").WebList("userLicenseReqForm:userTypeCombo")_;_script infofile_;_ZIP::ssf6.xml_;_

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebButton("Clear").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebButton("Clear").Click
End If

Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").WaitProperty "visible",1

item = Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").GetROProperty("value")

If item = "#0" Then
	Reporter.ReportEvent micPass,"Clear Button","Screen gets reset after clicking the reset button"
	Else
	Reporter.ReportEvent micFail,"Clear Button","Screen does not get reset after clicking the reset button"
End If

'Enter state name and license and user type

Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").WaitProperty "visible",1

Wait(2)

Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").Select "#1"

Wait(2)

Browser("EPS - User License Requirement").Page("EPS - User License Requirement").WebList("userLicenseReqForm:licenseType").Select strlicensetypename

Wait(2)

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebList("userLicenseReqForm:userTypeCombo").Select DataTable.Value("UserType","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS - User License Requirement 2").WebList("userLicenseReqForm:userTypeCombo")_;_script infofile_;_ZIP::ssf6.xml_;_

Wait(1)

'Click on Add

If Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Click
End If

Wait(3)

'Add duplicate record

Browser("Enterprise Control Center").Page("EPS - User License Requirement").WebList("userLicenseReqForm:stateCombo").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebList("userLicenseReqForm:licenseType").Select "RPh"

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebList("userLicenseReqForm:licenseType").Select strlicensetypename

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebList("userLicenseReqForm:userTypeCombo").Select DataTable.Value("UserType","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS - User License Requirement 2").WebList("userLicenseReqForm:userTypeCombo")_;_script infofile_;_ZIP::ssf6.xml_;_

'Click on Add

If Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("EPS - User License Requirement").WebButton("Add").Click
End If

'Checkpoint for duplicate record

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").Check CheckPoint("EPS - User License Requirement") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS - User License Requirement_2")_;_script infofile_;_ZIP::ssf16.xml_;_

'Click on Deactivate

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").Link("Deactivate").Click
End If

'Navigate to Administration>User Security>User Licenses>EPS User License Type

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebElement("header-form:j_id107:anchor").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebElement("header-form:j_id107:anchor").Click
End If

'Search License type

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebEdit("licenseTypes:licenseTypeSearch").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebEdit("licenseTypes:licenseTypeSearch").Set strlicensetypename

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebButton("Search").Click
End If

'Click on Deactivate

Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").Link("Deactivate").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").Link("Deactivate").Click
End If

'Navigate back to Home Screen

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebElement("WebTable").Click
End If


'Navigate to Administration>User Security>User Licenses>EPS User License Requirements

If Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebElement("header-form:j_id109:anchor").Exist(15) Then
Browser("EPS - User License Requirement_2").Page("Enterprise Control Center").WebElement("header-form:j_id109:anchor").Click
End If

'Select State for which you need to edit the license

rowval = Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebTable("User License State Requirements").GetROProperty("rows")

For i = 0 To rowval Step 1
	
	celdt = Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebTable("User License State Requirements").GetCellData(i,"1")
	
	If Instr(celdt,strstate)>0 Then
	
	Set tem = Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebTable("User License State Requirements").ChildItem(i,2,"Link",0)	
	tem.Click
	Exit For
	End If
Next

'Check Show Deactivated checkbox

Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebCheckBox("userLicenseReqForm:showDeactivateCh").WaitProperty "visible",1


Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebCheckBox("userLicenseReqForm:showDeactivateCh").Set "ON" @@ hightlight id_;_Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebCheckBox("userLicenseReqForm:showDeactivateCh")_;_script infofile_;_ZIP::ssf25.xml_;_

'Click on Activate

If Browser("EPS - User License Requirement").Page("EPS - User License Requirements").Link("Activate").Exist(15) Then
Browser("EPS - User License Requirement").Page("EPS - User License Requirements").Link("Activate").Click
End If

Wait(3)

'Checkpoint after activating license

If Browser("EPS - User License Requirement_3").Page("EPS - User License Requirement").WebElement("*Must reactivate License").Exist(15) Then
Reporter.ReportEvent micPass,"Reactivate License","Displayed"

value = Browser("EPS - User License Requirement_3").Page("EPS - User License Requirement").WebElement("*Must reactivate License").GetROProperty("innerhtml")
strvalue = CStr(value)
Wait(2)

Reporter.ReportEvent micPass,"Reactivate license record for deactivated license type","Must reactivate License Type"

End If


'Click on Home button

If Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebElement("WebTable").Exist(15) Then
Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebElement("WebTable").Click
End If

'Close Internet Explorer

If Browser("EPS - User License Requirement_3").Exist(15) Then
Browser("EPS - User License Requirement_3").Close
End If



Reporter.ReportEvent micDone,"Steps","All steps for the test case have been completed" @@ hightlight id_;_Browser("Enterprise Control Center 2").Page("EPS User License Types").WebTable("Home")_;_script infofile_;_ZIP::ssf34.xml_;_
