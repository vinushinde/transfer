'....................................................................................................................................

'Test Name : AddNewUser1

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

'Import variables from sheet

licensename = DataTable.Value("LicenseTypeName","Global")
strlicensename = CStr(licensename)


'Navigate to Administration>User
 @@ hightlight id_;_10652435_;_script infofile_;_ZIP::ssf2.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_9267311_;_script infofile_;_ZIP::ssf3.xml_;_

'Search and select User by Employee ID

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set DataTable.Value("User2EmployeeID","Global")

Wait(1)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

Wait(2)

'Click on Add License button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Enter all mandatory information for adding a new license


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select "TX"

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("User2LicenseNumber1","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select DataTable.Value("LicenseTypeName","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set "07-21-2018" @@ hightlight id_;_6532551_;_script infofile_;_ZIP::ssf27.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Type micTab

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

'Click on Add License button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Enter all mandatory information for adding a new license


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select "TX"

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("User2LicenseNumber2","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select "RPh" @@ hightlight id_;_22422651_;_script infofile_;_ZIP::ssf32.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set "07-21-2018" @@ hightlight id_;_6532551_;_script infofile_;_ZIP::ssf27.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Type micTab

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

'Click on Save

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

'Enter User Authentication
 @@ hightlight id_;_29899348_;_script infofile_;_ZIP::ssf37.xml_;_
 
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("UserName1","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password1","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Reporter.ReportEvent micDone,"Add Licenses User 2","New licenses have been addded"
