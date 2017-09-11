'...................................................................................................................................

'Test Name : Validations_C150046

'Test Description : This test will perform all the validations for this test case

'Author : Kashish Ambwani

'Date Modified : 17th June 2017

'...................................................................................................................................

'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

'Navigate to Administration>User Security>Role Settings>EPS Role Settings

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id105:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id105:anchor")_;_script infofile_;_ZIP::ssf1.xml_;_

'Search and select User required EPS Role

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").WaitProperty "visible",1


Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch").Set DataTable.Value("EPSRole","Global") @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebEdit("searchForm:roleSearch")_;_script infofile_;_ZIP::ssf2.xml_;_

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Exist(10) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebButton("Search").Click

Reporter.ReportEvent micPass,"Search EPS Role","User is able to search for EPS Role"
Else
Reporter.ReportEvent micFail,"Search EPS Role","User is unable to search for EPS Role"

End If

Wait(2)

'Click on Add

If Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Add").Click

Reporter.ReportEvent micPass,"Add Role","User is able to add role as supplemental role"
Else
Reporter.ReportEvent micFail,"Add Role","User is unable to add role as supplemental role"

End If

Wait(3)

roledescription = DataTable.Value("EPSRole","Global")

'Query to find status of supplemental role flag after adding role as supplemental role

supplementaladd = "select (ECC.EPS_ROLE.ALLOW_AS_SUPPLEMENTAL) ALLOW from ECC.EPS_ROLE where ECC.EPS_ROLE.DESCRIPTION = '"+ roledescription +"' and ECC.EPS_ROLE.ALLOW_AS_SUPPLEMENTAL IS NOT NUll"

'Execute Query
Set rcEPSRecordSet =  EccDBObject.Execute(supplementaladd)

'assigning value to supplemental role status

roleadd = rcEPSRecordSet.Fields("ALLOW")
strroleadd = CStr(roleadd)

If roleadd = "Y" Then
	Reporter.ReportEvent micPass,"Allow as Supplemental Role - DB Validation","Y"
	Else
	Reporter.ReportEvent micFail,"Allow as Supplemental Role - DB Validation",""+ strroleadd +""
End If

Wait(3)

'Remove role as supplemental role

If Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Remove").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Remove").Click

Reporter.ReportEvent micPass,"Remove Role","User  is able to remove role"
Else
Reporter.ReportEvent micFail,"Remove Role","User  is unable to remove role"

End If

Wait(3)

'Query to find status of supplemental role flag after removing role as supplemental role

supplementalremove = "select (ECC.EPS_ROLE.ALLOW_AS_SUPPLEMENTAL) ALLOW from ECC.EPS_ROLE where ECC.EPS_ROLE.DESCRIPTION = '"+ roledescription +"' and ECC.EPS_ROLE.ALLOW_AS_SUPPLEMENTAL IS NOT NUll"

'Execute Query
Set rcEPSRecordSet =  EccDBObject.Execute(supplementalremove)

'assigning value to supplemental role status

roleremove = rcEPSRecordSet.Fields("ALLOW")
strroleremove = CStr(roleremove)

If roleremove = "N" Then
	Reporter.ReportEvent micPass,"Allow as supplemental role - remove DB","N"
	Else
	Reporter.ReportEvent micFail,"Allow as supplemental role - remove DB",""+ strroleremove +""
End If

'Click on Home

If Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("WebTable").Exist(15) Then
Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("WebTable").Click
End If


Reporter.ReportEvent micDone,"Validations - ECC","All validations have been covered"
