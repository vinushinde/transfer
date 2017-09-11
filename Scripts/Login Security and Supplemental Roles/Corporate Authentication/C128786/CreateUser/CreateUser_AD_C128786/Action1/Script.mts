'...........................................................................................................................

'Test Name : CreateUser_AD_C128786

'Test Description : This test will create a new Corporate Authentication User

'Author : Kashish Ambwani

'Date Modified : 17 July 2017

'...........................................................................................................................

'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

'Import sheet

DataTable.ImportSheet "E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\Corporate Authentication\C128786\C128786.xls",1,"Global"

'Import variables from sheet

csvfilepath = "E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\Corporate Authentication\config\users.csv"
jarfilepath = "E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\Corporate Authentication\config\ADTestTool.jar"
employeenum = DataTable.Value("Employee_number","Global")
userlogin = DataTable.Value("Login","Global")
userpassword = DataTable.Value("Password","Global")
userfirstname = DataTable.Value("FirstName","Global")
usermiddlename = DataTable.Value("MiddleName","Global")
userlastname = DataTable.Value("LastName","Global")
userinitials = DataTable.Value("Initials","Global")
userdeactivated = DataTable.Value("Deactivated","Global")
useraccountcontrol = DataTable.Value("UserAccountControl","Global")
usergroup1 = DataTable.Value("Group1","Global")
usergroup2 = DataTable.Value("Group2","Global")


nxtline = ""+ employeenum +","+ userlogin +","+ userpassword +","+ userfirstname +","+ usermiddlename +","+ userlastname +","+ userinitials +","+ userdeactivated +","+ useraccountcontrol +","+ usergroup1 +","+ usergroup2 +""


'Call actions here

Call UpdateCSV(csvfilepath,nxtline)
Wait(Iteration_Wait)
Call Runbat()

Reporter.ReportEvent micDone,"Create user","New User for AD has been created"





Call Runbat()




'
'Set WshShell = CreateObject("WScript.Shell")
'WshShell.Run "C:\testfol\test.bat" ,1, True



Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """E:\Sprint\Scripts\Costco\Login Security and Supplemental Roles\Corporate Authentication\config\test.bat""" ,1, True


'
'dim shell
'set shell=createobject("wscript.shell")
'shell.run "C:\testfol\test.bat", 1, True

'Wait(5)
