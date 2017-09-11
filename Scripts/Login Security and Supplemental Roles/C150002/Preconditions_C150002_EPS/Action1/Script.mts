'.......................................................................................................................................

'Test name : Preconditions_C150002_EPS

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
	releasebase = WshSysEnv("vRELEASE")
	
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

'Store data in variables

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

userempid = RandomUserType(5)
userinitailsnew = RandomUserType(2)
usernewgroup = DataTable.Value("User_Group","Global")
usernewpassword = DataTable.Value("User_Password","Global")
usernewtype = DataTable.Value("User_UserType","Global")

WshSysEnv("UserEmpID") = userempid

'Call actions

CreateNewUserEPS_New userempid,userempid,userempid,userinitailsnew,usernewgroup,userempid,usernewpassword,usernewtype,username,password
Wait(Iteration_Wait)

PingECC() @@ hightlight id_;_18023805_;_script infofile_;_ZIP::ssf90.xml_;_


'Add supplemental roles on newly added user

'Navigate to Administration>User

JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_31377479_;_script infofile_;_ZIP::ssf24.xml_;_

'Search user by employee ID

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set userempid

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

Wait(2)

'Add 1st supplemental role on User which will expire 5 days from now

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If

'Select 1st role from dropdown

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").Select DataTable.Value("SupplementalRole1","Global") @@ hightlight id_;_2890882_;_script infofile_;_ZIP::ssf28.xml_;_
 @@ hightlight id_;_15933685_;_script infofile_;_ZIP::ssf29.xml_;_
 
If CStr(releasebase) < "2609" Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Type micTab
Else
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Type micTab
End If 


vardate = DateAdd("d",5,Now())
a = split(vardate," ")
b = split(a(0),"-")
role1date = b(1)&"-"&b(0)&"-"&b(2)

If CStr(releasebase) < "2609" Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set role1date @@ hightlight id_;_706643_;_script infofile_;_ZIP::ssf36.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Type micTab

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set role1date

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Type micTab

End If


If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
End If

Wait(2)

'Add second role on user


If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If

'Select 2nd role from dropdown

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").Select DataTable.Value("SupplementalRole2","Global") @@ hightlight id_;_2890882_;_script infofile_;_ZIP::ssf28.xml_;_
 @@ hightlight id_;_15933685_;_script infofile_;_ZIP::ssf29.xml_;_



If CStr(releasebase) < "2609" Then

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Type micTab

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set role1date

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Type micTab

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Type micTab

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set role1date

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Type micTab

End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
End If

Wait(2)

'Add third role on user


If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If

'Select 3rd role from dropdown

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").Select DataTable.Value("SupplementalRole3","Global") @@ hightlight id_;_2890882_;_script infofile_;_ZIP::ssf28.xml_;_
 @@ hightlight id_;_15933685_;_script infofile_;_ZIP::ssf29.xml_;_
If CStr(releasebase) < "2609" Then

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Type micTab

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set role1date

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Type micTab

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Type micTab

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set role1date

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Type micTab

End If 
 


If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
End If

Wait(2)

'Add 4 licenses on User

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Add 1st license on User

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("State","Global") @@ hightlight id_;_9309438_;_script infofile_;_ZIP::ssf50.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf51.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("LicenseNumber1","Global") @@ hightlight id_;_8306170_;_script infofile_;_ZIP::ssf52.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select "RPh" @@ hightlight id_;_12999061_;_script infofile_;_ZIP::ssf53.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set role1date

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Add second license on User

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("State","Global") @@ hightlight id_;_9309438_;_script infofile_;_ZIP::ssf50.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf51.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("LicenseNumber2","Global") @@ hightlight id_;_8306170_;_script infofile_;_ZIP::ssf52.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select "RPh" @@ hightlight id_;_12999061_;_script infofile_;_ZIP::ssf53.xml_;_

licensedatestart = DateAdd("d",-5,Now())
c = split(licensedatestart," ")
d = split(c(0),"-")
license2date = d(1)&"-"&d(0)&"-"&d(2)

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set license2date

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If @@ hightlight id_;_14767400_;_script infofile_;_ZIP::ssf61.xml_;_

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Add 3rd License on User

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("State","Global") @@ hightlight id_;_9309438_;_script infofile_;_ZIP::ssf50.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf51.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("LicenseNumber3","Global") @@ hightlight id_;_8306170_;_script infofile_;_ZIP::ssf52.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select "RPh" @@ hightlight id_;_12999061_;_script infofile_;_ZIP::ssf53.xml_;_

licensedatestart3 = DateAdd("d",-30,Now())
e = split(licensedatestart3," ")
f = split(e(0),"-")
license3date = f(1)&"-"&f(0)&"-"&f(2)


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set license3date

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add License").Click
End If

'Add 4th license on User


JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("State").Select DataTable.Value("State","Global") @@ hightlight id_;_9309438_;_script infofile_;_ZIP::ssf50.xml_;_
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf51.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("License Number").Set DataTable.Value("LicenseNumber4","Global") @@ hightlight id_;_8306170_;_script infofile_;_ZIP::ssf52.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaList("License Type").Select "RPh" @@ hightlight id_;_12999061_;_script infofile_;_ZIP::ssf53.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaEdit("Expiration Date").Set ""

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Licenses").JavaButton("Save").Click
End If @@ hightlight id_;_12703603_;_script infofile_;_ZIP::ssf73.xml_;_

Wait(1) @@ hightlight id_;_14916470_;_script infofile_;_ZIP::ssf72.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("UserName","Global") @@ hightlight id_;_5624543_;_script infofile_;_ZIP::ssf74.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("Password","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If


Wait(2)

'Expire 2nd role on User by removing role from ECC

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


LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password

Wait(3)


'Navigate to Administration>User Security>Role Settings>EPS Role Settings

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebElement("header-form:j_id105:anchor").Click

'Remove second supplemental role

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebEdit("searchForm:roleSearch").WaitProperty "visible",1

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebEdit("searchForm:roleSearch").Set DataTable.Value("SupplementalRole2","Global")

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

Wait(3)

'Expire 3rd role from DB by changing expiration date

strrole3 = CStr(DataTable.Value("SupplementalRole3","Global"))

'Query to find id of role

RoleID = "select ("+ vSchema +"ROLES.ID) ROLEID from "+ vSchema +"ROLES where "+ vSchema +"ROLES.DESCRIPTION = '"+ strrole3 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(RoleID)

'assigning value to supplemental role id

roleid = rcEPSRecordSet.Fields("ROLEID")
strroleid = CStr(roleid)

Wait(3)

'Query to expire role from DB

Expirerole3 = "update "+ vSchema +"Supplemental_role_link set "+ vSchema +"Supplemental_role_link.ENDING_DATE= sysdate -30 where "+ vSchema +"SUPPLEMENTAL_ROLE_LINK.ROLE_ID = '"+ strroleid +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Expirerole3)

PingECC()


Reporter.ReportEvent micDone,"Preconditions - EPS","All preconditions have been set on EPS"
