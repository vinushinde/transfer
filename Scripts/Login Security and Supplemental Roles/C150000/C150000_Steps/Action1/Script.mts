'..........................................................................................................................................

'Test Name : C150000_Steps

'..........................................................................................................................................

'Regression testing connection:

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

vServerip = WshSysEnv("server")
vDbuser = WshSysEnv("dbeccuser")
vDbpwd = WshSysEnv("dbpassword")

Set EccDBObject = CreateObject("ADODB.Connection")

Call ECC_SPDBConnection(EccDBObject,vServerip,vDbuser,vDbpwd)

'Call Steps from here


