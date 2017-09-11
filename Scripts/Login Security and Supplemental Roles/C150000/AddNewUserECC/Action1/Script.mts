'.................................................................................................................................

'Test Name : AddNewUserECC

'.................................................................................................................................

'Navigate to Administration>User Security>user settings

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id98:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id98:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Click on Add User

Browser("Enterprise Control Center").Page("Users").WebButton("Add User").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("Users").WebButton("Add User").Exist(15) Then
Browser("Enterprise Control Center").Page("Users").WebButton("Add User").Click
End If

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:employeeId").WaitProperty "visible",1


Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:employeeId").Set DataTable.Value("User2EmployeeID","Global")

Wait(1)

 @@ hightlight id_;_Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:employeeId")_;_script infofile_;_ZIP::ssf4.xml_;_
Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:lastName").Set DataTable.Value("User2lastname","Global")

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:firstName").Set DataTable.Value("User2FirstName","Global")

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:userInitials").Set DataTable.Value("User2Initials","Global")

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebList("editUserForm:userGroup").Select "DoAllGroup" @@ hightlight id_;_Browser("Enterprise Control Center").Page("Users").WebList("editUserForm:userGroup")_;_script infofile_;_ZIP::ssf8.xml_;_

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:userId").Set DataTable.Value("User2userID","Global")

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:reEnterLogonId").Set DataTable.Value("User2LogonID","Global")

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:password").Set DataTable.Value("User2Pasword","Global")

Wait(1)

Browser("Enterprise Control Center").Page("Users").WebEdit("editUserForm:reEnterPassword").Set DataTable.Value("User2Pasword","Global")

'Browser("Enterprise Control Center").Page("Users").WebElement("editUserForm").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Users").WebElement("editUserForm")_;_script infofile_;_ZIP::ssf13.xml_;_
Wait(1)

Browser("Enterprise Control Center").Page("Users").WebList("editUserForm:userType").Select "RPh" @@ hightlight id_;_Browser("Enterprise Control Center").Page("Users").WebList("editUserForm:userType")_;_script infofile_;_ZIP::ssf14.xml_;_
 @@ hightlight id_;_Browser("Enterprise Control Center").Page("Add/Edit User").WebButton("Save")_;_script infofile_;_ZIP::ssf24.xml_;_
 Wait(1)
 
 'Click on Save
 
 If Browser("Enterprise Control Center").Page("Users").WebButton("Save").Exist(15) Then
 	Browser("Enterprise Control Center").Page("Users").WebButton("Save").Click
 End If
 
 Wait(1)
 
'Click on Home button

If Browser("Enterprise Control Center").Page("Users").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center").Page("Users").WebElement("WebTable").Click
End If

