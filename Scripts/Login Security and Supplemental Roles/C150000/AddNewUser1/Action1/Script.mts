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

'Query to find if license type has been added to license type table

licensetype = "select count("+ veccSchema +"LICENSE_TYPE.LICENSE_TYPE) COUNT from "+ veccSchema +"LICENSE_TYPE where "+ veccSchema +"LICENSE_TYPE.LICENSE_TYPE = '"+ strlicensename +"'"

'Execute Query
Set rcEPSRecordSet =  EccDBObject.Execute(licensetype)

'assigning value to count

licensecount = rcEPSRecordSet.Fields("COUNT")
intlicensecount = CInt(licensecount)

If intlicensecount>0 Then
	Reporter.ReportEvent micPass,"License Type Entry - DB","License type is present in ECC DB"
	Else
	Reporter.ReportEvent micFail,"License Type Entry - DB","License type is not present in ECC DB"
End If

Wait(2)

PingECC()

'Navigate to Administration>User
 @@ hightlight id_;_10652435_;_script infofile_;_ZIP::ssf2.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_9267311_;_script infofile_;_ZIP::ssf3.xml_;_

'Search and select User by Employee ID

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set DataTable.Value("EmployeeID","Global")

Wait(1)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

'Click OK on message dialog

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Message").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Message").JavaButton("OK").Click
End If

'Click on Add new User

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add New User").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add New User").Click
End If

Wait(2)

'Enter all mandatory information for adding new user

JavaWindow("Enterprise Pharmacy System").JavaEdit("Last Name").Set DataTable.Value("User1LastName","Global")
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf9.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").Set DataTable.Value("User1FirstName","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Initials").Set DataTable.Value("User1Initials","Global")

'Search for group

If JavaWindow("Enterprise Pharmacy System").JavaButton("find").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Group Search").JavaEdit("Group Name").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Group Search").JavaEdit("Group Name").Set "do"
 @@ hightlight id_;_12505368_;_script infofile_;_ZIP::ssf14.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Group Search").JavaEdit("Group Name").Type micReturn

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Group Search").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Group Search").JavaButton("Select").Click
End If

'Enter User login Code

JavaWindow("Enterprise Pharmacy System").JavaEdit("User Login Code").Set DataTable.Value("UserName1","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password1","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Confirm Password").Set DataTable.Value("Password1","Global")

JavaWindow("Enterprise Pharmacy System").JavaEdit("Employee ID Number").Set DataTable.Value("EmployeeID","Global")

'Click on Add License button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Enter all mandatory information for adding a new license


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("State","Global") @@ hightlight id_;_12263310_;_script infofile_;_ZIP::ssf21.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("User1LicenseNumber1","Global")

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


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("State","Global") @@ hightlight id_;_12263310_;_script infofile_;_ZIP::ssf21.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("User1LicenseNumber2","Global")

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
 
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("UserName","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Reporter.ReportEvent micDone,"Add New user 1","New User has been added"
