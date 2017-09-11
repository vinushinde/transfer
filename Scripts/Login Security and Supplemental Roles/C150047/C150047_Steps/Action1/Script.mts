'....................................................................................................................................

'Test Name : C150047_Steps

'....................................................................................................................................


'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

'Call variables from sheet

username1 = DataTable.Value("UserName1","Global")
password1 = DataTable.Value("Password1","Global")


'Navigate to Administration>User Security>User License>EPS User License Requirements

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id108:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id108:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Click on Add New License Type

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Add License Type").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Add License Type").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Add License Type")_;_script infofile_;_ZIP::ssf2.xml_;_
 @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types")_;_script infofile_;_ZIP::ssf3.xml_;_
'Enter details for the license type

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseType").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseType").Set DataTable.Value("LicenseTypeName","Global")

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseDescr").Set DataTable.Value("LicenseTypeDescription","Global")

'Click on Save

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save").Click
End If

'Search for license type

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch").Set DataTable.Value("LicenseTypeName","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch")_;_script infofile_;_ZIP::ssf7.xml_;_

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search").Click
End If

Wait(2)

'Click on Edit

If Browser("Enterprise Control Center").Page("EPS User License Types").Link("Edit").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").Link("Edit").Click
End If

'Change Description

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseDescr").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseDescr").Set DataTable.Value("LicenseTypeDescription2","Global")

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save").Click
End If


'Search for license type

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch").WaitProperty "visible",1

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch").Set DataTable.Value("LicenseTypeName","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch")_;_script infofile_;_ZIP::ssf7.xml_;_

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search").Click
End If @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search")_;_script infofile_;_ZIP::ssf13.xml_;_

'Click on Deactivate

Browser("Enterprise Control Center").Page("EPS User License Types").Link("Deactivate").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("EPS User License Types").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").Link("Deactivate").Click
End If

'Click on Activate

Browser("Enterprise Control Center").Page("EPS User License Types").Link("Activate").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("EPS User License Types").Link("Activate").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").Link("Activate").Click
End If

'Click on Add New License Type

Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Add License Type").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Add License Type").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Add License Type").Click
End If

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseType").Set DataTable.Value("LicenseTypeName","Global")

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("addEditUserLicenseForm:licenseDescr").Set DataTable.Value("LicenseTypeDescription2","Global")

'Click on Save

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save").Click
End If @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Save")_;_script infofile_;_ZIP::ssf20.xml_;_

'Checkpoint for Duplicate License Type(s)

Browser("Enterprise Control Center").Page("EPS User License Types").Check CheckPoint("EPS User License Types") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types")_;_script infofile_;_ZIP::ssf21.xml_;_

'Click on Cancel

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Cancel").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Cancel").Click
End If

'Search for RPh/CPht License Type

Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch").WaitProperty "visible",1


Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch").Set "RPh" @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types").WebEdit("licenseTypes:licenseTypeSearch")_;_script infofile_;_ZIP::ssf23.xml_;_

If Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebButton("Search").Click
End If

'Click on deactivate

If Browser("Enterprise Control Center").Page("EPS User License Types").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").Link("Deactivate").Click
End If

Wait(2)

'Checkpoint after trying to deactivate RPh/CPht License Type

Browser("Enterprise Control Center").Page("EPS User License Types").Check CheckPoint("EPS User License Types_2") @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS User License Types")_;_script infofile_;_ZIP::ssf27.xml_;_

'Click on Home button

If Browser("Enterprise Control Center").Page("EPS User License Types").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS User License Types").WebElement("WebTable").Click
End If

'Navigate to Administration>User Security>User License>EPS User License Requirements

Browser("Enterprise Control Center").Page("Enterprise Control Center_2").WebElement("header-form:j_id109:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center 2").WebElement("header-form:j id109:anchor")_;_script infofile_;_ZIP::ssf29.xml_;_


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

'Select state and license type

Browser("EPS - User License Requirement").Page("EPS - User License Requirement").WebList("userLicenseReqForm:licenseType").Select DataTable.Value("LicenseTypeName","Global")

Wait(2)

'Select User Type from dropdown

Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebList("userLicenseReqForm:userTypeCombo").Select "RPh" @@ hightlight id_;_Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebList("userLicenseReqForm:userTypeCombo")_;_script infofile_;_ZIP::ssf93.xml_;_

'Click on Add

If Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebButton("Add").Exist(15) Then
Browser("EPS - User License Requirement").Page("EPS - User License Requirements").WebButton("Add").Click
End If


Reporter.ReportEvent micPass,"License Type dropdown","New License is present in dropdown"

'Click on Home button

If Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - User License Requirement_2").WebElement("WebTable").Click
End If

'..............................................ECC Settings are complete.........................................................................

Wait(5)

''..............................................EPS part..........................................................................................

'Click on Logout

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Login to EPS with User1

EPSLogin username1,password1

'Ping ECC

PingECC()

'Check if new license is added under user license on EPS

'Navigate to Administration>User

JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_20427694_;_script infofile_;_ZIP::ssf41.xml_;_

'Search and select user

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set DataTable.Value("EmployeeID1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

Wait(2)

'Add custom license on user 1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").WaitProperty "visible",1


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("LicenseState","Global") @@ hightlight id_;_30797465_;_script infofile_;_ZIP::ssf45.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("LicenseNumber1","Global") @@ hightlight id_;_21155980_;_script infofile_;_ZIP::ssf46.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select DataTable.Value("LicenseTypeName","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set DataTable.Value("ExpirationDate1","Global") @@ hightlight id_;_26710249_;_script infofile_;_ZIP::ssf50.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("UserName1","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

Wait(2)

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Wait(2)

'.........................................................ECC part...............................................................................

'Navigate to Administration>User Security>User Licenses>EPS License Type

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebElement("header-form:j_id108:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center 2").Page("Enterprise Control Center").WebElement("header-form:j id108:anchor")_;_script infofile_;_ZIP::ssf57.xml_;_

'Search for license type

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebEdit("licenseTypes:licenseTypeSearch").WaitProperty "visible",1

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebEdit("licenseTypes:licenseTypeSearch").Set DataTable.Value("LicenseTypeName","Global")

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebButton("Search").Click
End If

'Click on Deactivate

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").Link("Deactivate").WaitProperty "visible",1

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center").Link("Deactivate").Click
End If

Wait(2)

'Click on Home button

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebElement("WebTable").Click
End If


'.........................................................EPS Part...............................................................................

UserLicesnseExpiration("Enabled")

'Ping ECC

PingECC()

'Change authentication mode on EPS

AuthenticationModeEPS(DataTable.Value("AuthenticationMode2","Global")) @@ hightlight id_;_32997372_;_script infofile_;_ZIP::ssf70.xml_;_

'Logout and login again with user 1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Login to EPS with user1

EPSLogin username1,password1 

'Checkpoint for Uer Login Alert

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.(st)") @@ hightlight id_;_20860993_;_script infofile_;_ZIP::ssf75.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If

''.................................................ECC Part.....................................................................................

'Activate User License Type

'Navigate to Administration>User Security>User Licenses>EPS License Type

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebElement("header-form:j_id108:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center 2").Page("Enterprise Control Center").WebElement("header-form:j id108:anchor")_;_script infofile_;_ZIP::ssf77.xml_;_

'Search for License type

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebEdit("licenseTypes:licenseTypeSearch").WaitProperty "visible",1

Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebEdit("licenseTypes:licenseTypeSearch").Set DataTable.Value("LicenseTypeName","Global")

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebButton("Search").Click
End If

Wait(2)

'Click on Activate

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center").Link("Activate").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center").Link("Activate").Click
End If

Wait(3)

'Click on Home Button

If Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center_2").Page("Enterprise Control Center").WebElement("WebTable").Click
End If
 @@ hightlight id_;_29642872_;_script infofile_;_ZIP::ssf91.xml_;_
Wait(2)

'.........................................................................................................................................


'.........................................................EPS Part..........................................................................


'Logout and login again with user 1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Login to EPS with user1

EPSLogin username1,password1 

'Checkpoint for Uer Login Alert

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.(st)") @@ hightlight id_;_20860993_;_script infofile_;_ZIP::ssf75.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If
 
Wait(2) 

'Logout and login again with user 1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Login with User 2

username2 = DataTable.Value("UserName2","Global")
password2 = DataTable.Value("Password2","Global")

EPSLogin username2,password2

'Login with akumar User

'Logout and login again with user 1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

'Login with User 2

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

EPSLogin username,password


'Clsoe Internet Explorer
 @@ hightlight id_;_Browser("EPS - User License Requirement").Page("EPS - User License Requirements")_;_script infofile_;_ZIP::ssf102.xml_;_
 If Browser("EPS - User License Requirement").Exist(15) Then
 Browser("EPS - User License Requirement").Close
 End If

 
 Reporter.ReportEvent micDone,"Steps","All steps have been executed"
