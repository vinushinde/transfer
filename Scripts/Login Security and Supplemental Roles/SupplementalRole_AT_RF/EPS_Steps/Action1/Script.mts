'...........................................................................................................................................

'Test Name : C150045_Steps

'...........................................................................................................................................

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


'Import variables from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

password1 = DataTable.Value("User_Password","Global")

username1 = RandomUserType(5)
usernewinitials = RandomUserType(3)
usernewgroup = DataTable.Value("User_Group","Global")
usernewtype = DataTable.Value("User_UserType","Global")


'Create New User

CreateNewUserEPS_New username1,username1,username1,usernewinitials,usernewgroup,username1,password1,usernewtype,username,password

'Ping ECC

PingECC()

'Click on logout

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If


'Login into EPS

EPSLogin username1,password1

'Handle Expired Dialog pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf51.xml_;_

'Navigate to Tools>Inventory Management

If JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Inventory Management").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Inventory Management").Select
End If

'......................................................Inventory Management Part..................................................................


If Browser("Certificate Error: Navigation_3").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(15) Then
Browser("Certificate Error: Navigation_3").Page("Certificate Error: Navigation").Link("Continue to this website").Click
End If

If Browser("Certificate Error: Navigation_3").Page("Certificate Error: Navigation_2").Link("Continue to this website").Exist(15) Then
Browser("Certificate Error: Navigation_3").Page("Certificate Error: Navigation_2").Link("Continue to this website").Click
End If



'Login to inventory management site 

Wait(5)

If Browser("Certificate Error: Navigation").Page("Rx.com Enterprise Pharmacy").WebEdit("j_username").Exist(10) Then
Browser("Certificate Error: Navigation").Page("Rx.com Enterprise Pharmacy").WebEdit("j_username").Set username1
End If

If Browser("Certificate Error: Navigation").Page("Rx.com Enterprise Pharmacy").WebEdit("j_password").Exist(10) Then
Browser("Certificate Error: Navigation").Page("Rx.com Enterprise Pharmacy").WebEdit("j_password").Set password1
End If


If Browser("Certificate Error: Navigation").Page("Rx.com Enterprise Pharmacy").WebButton("Login").Exist(15) Then
Browser("Certificate Error: Navigation").Page("Rx.com Enterprise Pharmacy").WebButton("Login").Click
End If

'Checkpoint for failed login

Wait(3)

Browser("Certificate Error: Navigation").Page("Enterprise Store-Based").Check CheckPoint("Enterprise Store-Based Mail Order - Access Denied") @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Store-Based")_;_script infofile_;_ZIP::ssf12.xml_;_

'Close Internet Explorer

Browser("Certificate Error: Navigation").Close

'.........................................................EPS Part........................................................................

'Navigate to Administration>User

JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select @@ hightlight id_;_19920028_;_script infofile_;_ZIP::ssf17.xml_;_

'Search for User

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1


JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set username1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

'............................................Add loop for all supplemental roles.........................................................



'********************

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If
Wait(4)
'JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").WaitProperty "visible",1

items = JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").GetROProperty("items count")
print"items value:"&items
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").Close

Dim temparr(50)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If
For x = 1 To items-1 Step 1
	temparr(x) = JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").GetItem(x)
Next

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").Close
For i = 1 To items-1 Step 1
print"i value:"&i	
If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If	

print"itemName value:"&itemName	

first1 = Date()
second1 = split(first1,"-")
today = second1(1)&"-"&second1(0)&"-"&second1(2)

Wait(1)

first2 = DateAdd("d",5,Date())
second2 = split(first2,"-")
expiration = second2(1)&"-"&second2(0)&"-"&second2(2)


Wait(5) @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf22.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental").Type temparr(i)
Wait(2)

If CStr(releasebase) < "2609" Then

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Set today
 @@ hightlight id_;_21771077_;_script infofile_;_ZIP::ssf29.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set expiration
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Type micTab

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Set today
 @@ hightlight id_;_21771077_;_script infofile_;_ZIP::ssf29.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set expiration
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Type micTab

End If

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
	
Next

'*****************************

Wait(3)


If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
	JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If


'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set username1 @@ hightlight id_;_24734978_;_script infofile_;_ZIP::ssf115.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set password1

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If


Wait(2)


'........................................................................................................................................

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

'..........................................Checkpoint for inventory management........................................................

If JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Inventory Management").Select
End If

'Certificate Error

If Browser("Certificate Error: Navigation_4").Page("Certificate Error: Navigation").Link("Continue to this website").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Certificate Error: Navigation").Link("Continue to this website").Click
End If

If Browser("Certificate Error: Navigation_4").Page("Certificate Error: Navigation_2").Link("Continue to this website").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Certificate Error: Navigation_2").Link("Continue to this website").Click
End If


'Enter Username and Password on Inventory Management login screen

If Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_username").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_username").Set username1
End If

If Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_password").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_password").Set password1
End If


If Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebButton("Login").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebButton("Login").Click
End If

'Click on Back

If Browser("Certificate Error: Navigation_4").Exist(15) Then
Browser("Certificate Error: Navigation_4").Back
End If

'Enter Username and Password on Inventory Management login screen

If Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_username").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_username").Set username1
End If

If Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_password").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebEdit("j_password").Set password1
End If


If Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebButton("Login").Exist(15) Then
Browser("Certificate Error: Navigation_4").Page("Rx.com Enterprise Pharmacy").WebButton("Login").Click
End If
 @@ hightlight id_;_Browser("Certificate Error: Navigation 4").Page("Rx.com Enterprise Pharmacy 2").WebButton("Login")_;_script infofile_;_ZIP::ssf119.xml_;_
'Checkpoint for Inventory Management Screen

Wait(3)

Browser("Certificate Error: Navigation_4").Page("EPS Store Dashboard -").Check CheckPoint("EPS Store Dashboard - Summary") @@ hightlight id_;_Browser("Certificate Error: Navigation_4").Page("EPS Store Dashboard -")_;_script infofile_;_ZIP::ssf50.xml_;_
Browser("Certificate Error: Navigation_4").Close

'..............................................EPS Part...................................................................................


'Query to find role ID of Will Call Role

RoleId = "Select (EPS2.ROLES.ID) ROLEID from EPS2.ROLES where EPS2.ROLES.DESCRIPTION = 'Will Call'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(RoleId)

'assigning value to will call role id

roleid = rcEPSRecordSet.Fields("ROLEID")
strroleid = CStr(roleid)

Wait(3)

'Query to update will call role ending date

endingdate = "Update EPS2.SUPPLEMENTAL_ROLE_LINK Set EPS2.SUPPLEMENTAL_ROLE_LINK.ENDING_DATE = TO_TIMESTAMP ('2017-03-03 11:11:11','YYYY-MM-DD HH:MI:SS') where EPS2.SUPPLEMENTAL_ROLE_LINK.ROLE_ID = '"+ strroleid +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(endingdate)


'Logout of client

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If


'Login into EPS

EPSLogin username1,password1

'Handle Expired Dialog pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If


'Checkpoint for Will Call button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Will Call").GetROProperty("enabled")=0 Then
	Reporter.ReportEvent micPass,"User - Will Call Role","User is unable to access Will Call"
	Else
	Reporter.ReportEvent micFail,"User - Will Call Role","User is unable to access Will Call"
End If

'.......................................................................................................................................

'.................................................Check on User for expired role........................................................

'Navigate to Administration>User

If JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaMenu("Administration").JavaMenu("User").Select
End If

'Search for User by employee ID

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaEdit("Select by Employee ID").Set username1

Wait(1)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Select").Click
End If

Wait(3)

'Check Show Expired Checkbox

JavaWindow("Enterprise Pharmacy System").JavaCheckBox("Show Expired").Set "ON" @@ hightlight id_;_5055305_;_script infofile_;_ZIP::ssf63.xml_;_

Wait(3)

'Select Will Call role

rowval = JavaWindow("Enterprise Pharmacy System").JavaTable("Supplemental Roles").GetROProperty("rows")

Wait(2)

For i = 0 To rowval-1 Step 1
	
	getcell = JavaWindow("Enterprise Pharmacy System").JavaTable("Supplemental Roles").GetCellData(i,"Role Description")
	
	If getcell = "Will Call" Then
		JavaWindow("Enterprise Pharmacy System").JavaTable("Supplemental Roles").SelectRow i
		Reporter.ReportEvent micPass,"Will Call Role","Expired"
		Exit For
End If
	
Next

Wait(2)

'Click on Edit Supplemental Role button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Supplemental Role").Click
End If

Wait(2)

'Checkpoint after trying to edit supplemental role

JavaWindow("Enterprise Pharmacy System").JavaDialog("Error").JavaStaticText("Expiration date for this").Check CheckPoint("Expiration date for this supplemental role has been exceeded. Please add a new record.(st)") @@ hightlight id_;_9970864_;_script infofile_;_ZIP::ssf121.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Error").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Error").JavaButton("OK").Click
End If

'Click on Add Supplemental Role button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Add Supplemental Role").Click
End If

'Select Will Call from supplemental role dropdown

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental_2").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaList("Available Supplemental_2").Select "Will Call" @@ hightlight id_;_4201469_;_script infofile_;_ZIP::ssf124.xml_;_

'Set effective date before 180 days

If CStr(releasebase) < "2609" Then

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Set "06-25-2015" @@ hightlight id_;_21644625_;_script infofile_;_ZIP::ssf129.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Type micTab

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set "07-25-2017" @@ hightlight id_;_4689307_;_script infofile_;_ZIP::ssf134.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Type micTab

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Set "06-25-2015"

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Type micTab

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set "07-25-2017"

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Type micTab

End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
End If

'Checkpoint for effective date

'JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaDialog("Error").JavaStaticText("Expiration date must be").Check CheckPoint("Expiration date must be less than or equal to 12-22-2017(st)_2") @@ hightlight id_;_2728497_;_script infofile_;_ZIP::ssf131.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaDialog("Error").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaDialog("Error").JavaButton("OK").Click
End If

Wait(2)

'Set expiration date more than 180 days

If CStr(releasebase) < "2609" Then

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set "07-08-2018" @@ hightlight id_;_4689307_;_script infofile_;_ZIP::ssf125.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Type micTab

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set "07-08-2018"

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Type micTab

End If


If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
End If

'Checkpoint for date 

'errorfound = JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaDialog("Error").JavaStaticText("Expiration date and time").GetROProperty("text")
'strerror = CStr(errorfound)
'
'Reporter.ReportEvent micPass,"Error",""+ strerror +""

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaDialog("Error").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaDialog("Error").JavaButton("OK").Click
End If

Wait(2)

'Set expiration and effective date

If CStr(releasebase) < "2609" Then

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date").Set today

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date").Set expiration

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Effective Date and Time").Set today

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaEdit("Expiration Date and Time").Set expiration

End If



If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Supplemental Role").JavaButton("Save").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

'Enter User Authentication

JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set username1 @@ hightlight id_;_1790343_;_script infofile_;_ZIP::ssf137.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set password1

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Wait(3)




'..........................................................................................................................................



'Click on Logout

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

EPSLogin username1,password1

'Checkpoint for Will Call button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Will Call").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micPass,"Will Call button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Will Call button","Disabled"
End If
 @@ hightlight id_;_28225854_;_script infofile_;_ZIP::ssf91.xml_;_
'Login with default user

' Click on Logout

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If

EPSLogin username,password

Reporter.ReportEvent micDone,"Steps","All steps have been executed" @@ hightlight id_;_33121334_;_script infofile_;_ZIP::ssf112.xml_;_
