'...............................................................................................................................................

'Test Name : Preconditions_ECC_MMN

'Date Modified : 9 August 2017
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

'Import data from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")

state = DataTable.Value("ECCImmunizationState","Global")
minimumage = DataTable.Value("Min_Age","Global")
maximumage = DataTable.Value("Max_Age","Global")
sharevalueecc = DataTable.Value("ECC_ShareImmunization","Global")

'Call actions here

LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
ImmunizationAge_ECC state,minimumage,maximumage,sharevalueecc
Wait(Iteration_Wait)
PingECC()

Reporter.ReportEvent micDone,"Preconditions ECC","All preconditions on ECC have been set" @@ hightlight id_;_Browser("EPS - Default Immunization").Page("EPS - Default Immunization")_;_script infofile_;_ZIP::ssf41.xml_;_
