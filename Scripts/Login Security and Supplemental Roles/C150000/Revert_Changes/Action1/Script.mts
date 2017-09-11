'................................................................................................................................

'Test Description : This test will revert all the changes made during the test

'................................................................................................................................


'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

Serverip = WshSysEnv("vServerip")
Dbeccuser = WshSysEnv("vDbeccUser")
Dbpwd = WshSysEnv("vDbpwd")
veccSchema = WshSysEnv("eccvSchema")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,Serverip,Dbeccuser,Dbpwd)

'Import data from sheet

licensename = DataTable.Value("LicenseTypeName","Global")
strlicensename = CStr(licensename)

Wait(2)

'Query to deactivate license

LicenseID = "select ("+ veccSchema +"LICENSE_TYPE.ID) ID from "+ veccSchema +"LICENSE_TYPE where "+ veccSchema +"LICENSE_TYPE.LICENSE_TYPE = '"+ strlicensename +"'"

'Execute Query
Set rcEPSRecordSet =  EccDBObject.Execute(LicenseID)

'assigning value to license ID

licenseid = rcEPSRecordSet.Fields("ID")
strid = CStr(licenseid)

Wait(2)

''Query to deactivate license type
'
'Deactivatetype = "update "+ veccSchema +"LICENSE_TYPE SET "+ veccSchema +"LICENSE_TYPE.DEACTIVATE_DATE = sysdate -2 where "+ veccSchema +"LICENSE_TYPE.LICENSE_TYPE = '"+ strlicensename +"'"
'
''Execute Query
'Set rcEPSRecordSet =  EccDBObject.Execute(Deactivate)

Wait(2)

'Query to deactivate license requirement

Deactivate = "update "+ veccSchema +"LICENSE_REQUIREMENT SET "+ veccSchema +"LICENSE_REQUIREMENT.DEACTIVATE_DATE = sysdate -2 where "+ veccSchema +"LICENSE_REQUIREMENT.LICENSE_TYPE_ID = '"+ strid +"'"

'Execute Query
Set rcEPSRecordSet =  EccDBObject.Execute(Deactivate)


Wait(2)

PingECC()


Reporter.ReportEvent micDone,"Revert Changes","All changes have been reverted"
