'...............................................................................................................................................

'Test Name : Verify that ECC user is able to add new custom EPS user type which gets pinged down to EPS.

'Test Description : ECC user with all rights and roles to ECC

'Date Modified : 1 August 2017
'...............................................................................................................................................

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

WshSysEnv ("WF") = "C144240"

'Import sheet

importDataSheet()

'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")


usertype = RandomUserType(10)
strusertype = CStr(usertype)
usertypedescription = RandomDescription(60)


'Call actions here

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AddNewUserType usertype,usertypedescription
Wait(Iteration_Wait)


'Navigate to Administration>User Security>User Licenses>EPS User License Requirements

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_5").WebElement("header-form:j_id110:anchor").Click @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Control Center 5").WebElement("header-form:j id110:anchor")_;_script infofile_;_ZIP::ssf9.xml_;_

'Wait until page is displayed

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_5").WebButton("Add New").WaitProperty "visible",1

'Select State for which you need to edit the license

rowval = Browser("Certificate Error: Navigation").Page("EPS - User License Requirements").WebTable("User License State Requirements").GetROProperty("rows")

For i = 0 To rowval Step 1
	
	celdt = Browser("Certificate Error: Navigation").Page("EPS - User License Requirements").WebTable("User License State Requirements").GetCellData(i,"1")
	
	If Instr(celdt,"TEXAS")>0 Then
	
	Set tem = Browser("Certificate Error: Navigation").Page("EPS - User License Requirements").WebTable("User License State Requirements").ChildItem(i,2,"Link",0)	
	tem.Click
	Exit For
	End If
Next

'Check if newly added User Type is present in User Type dropdown on State Requirement Screen

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_5").WebList("userLicenseReqForm:userTypeCombo").WaitProperty "visible",1

Wait(5)

itemval = Browser("Certificate Error: Navigation").Page("Enterprise Control Center_5").WebList("userLicenseReqForm:userTypeCombo").GetROProperty("Items Count")

For i = 1 To itemval-1 Step 1

itemname = Browser("Certificate Error: Navigation").Page("Enterprise Control Center_5").WebList("userLicenseReqForm:userTypeCombo").GetItem(i)

If itemname = strusertype Then
	Reporter.ReportEvent micPass,"State Requirement Screen - User Type dropdown","Newly added license is present in User Type dropdown on State Requirement Screen"
	
	Exit For
	
End If
	
Next

'Navigate to Home Screen

Browser("Enterprise Control Center_3").Page("Enterprise Control Center").WebElement("WebTable").Click @@ hightlight id_;_Browser("Enterprise Control Center 3").Page("Enterprise Control Center").WebElement("WebTable")_;_script infofile_;_ZIP::ssf11.xml_;_

Wait(Iteration_Wait)

PingECC()
 @@ hightlight id_;_28017381_;_script infofile_;_ZIP::ssf13.xml_;_
'Navigate to Administration>User

JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_31364309_;_script infofile_;_ZIP::ssf14.xml_;_

'Click on Add New User button

JavaWindow("Enterprise Pharmacy System").JavaButton("Add New User").WaitProperty "visible",1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add New User").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add New User").Click
End If

Wait(5)

'Check if newly added license is present in User Type dropdown on User screen in EPS

epsitemval = JavaWindow("Enterprise Pharmacy System").JavaList("User Type").GetROProperty("items count")

For j = 0 To epsitemval-1 Step 1
	
itemnameeps = 	JavaWindow("Enterprise Pharmacy System").JavaList("User Type").GetItem(j)

If itemnameeps = strusertype Then
	
	Reporter.ReportEvent micPass,"User Type Dropdown - EPS User","Newly added user type is present"
	
	Exit For
	
End If
	
Next
 @@ hightlight id_;_6333187_;_script infofile_;_ZIP::ssf16.xml_;_
'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Wait(Iteration_Wait)

'Navigate to EPS User License Type Screen on ECC

Browser("Enterprise Control Center_4").Page("Enterprise Control Center").WebElement("header-form:j_id112:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center 4").Page("Enterprise Control Center").WebElement("header-form:j id112:anchor")_;_script infofile_;_ZIP::ssf19.xml_;_

'Search for User Type

Browser("Enterprise Control Center_4").Page("Enterprise Control Center").WebEdit("userTypes:userTypeSearch").WaitProperty "visible",1

Browser("Enterprise Control Center_4").Page("Enterprise Control Center").WebEdit("userTypes:userTypeSearch").Set usertype

If Browser("Enterprise Control Center_4").Page("Enterprise Control Center").WebButton("Search").Exist(15) Then
Browser("Enterprise Control Center_4").Page("Enterprise Control Center").WebButton("Search").Click
End If

Wait(5)

'Click on Deactivate

Browser("Enterprise Control Center_4").Page("Enterprise Control Center").Link("Deactivate").WaitProperty "visible",1

If Browser("Enterprise Control Center_4").Page("Enterprise Control Center").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center_4").Page("Enterprise Control Center").Link("Deactivate").Click
End If

'Navigate to Home screen

Browser("Enterprise Control Center_4").Page("Enterprise Control Center").WebElement("WebTable").Click @@ hightlight id_;_Browser("Enterprise Control Center 4").Page("Enterprise Control Center").WebElement("WebTable")_;_script infofile_;_ZIP::ssf23.xml_;_

Wait(Iteration_Wait)

CloseIECertificate()

Wait(Iteration_Wait)

PingECC()

Reporter.ReportEvent micDone,"C144240","Test case has been executed" @@ hightlight id_;_Browser("Enterprise Control Center 3").Page("EPS User Types 2").WebElement("WebTable")_;_script infofile_;_ZIP::ssf8.xml_;_
