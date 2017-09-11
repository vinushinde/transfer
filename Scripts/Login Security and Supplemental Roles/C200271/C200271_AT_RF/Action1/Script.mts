'............................................................................................................................

'Test Name : Verify License/Supplemental Role/Password about to expire and Expired data FIles PopUp to Display on secondary client when Authentication Mode is 2 and when client inactive session time out is '60'

'TestRail ID : C200271

'Author : Kashish Ambwani

'Date Modified : 1 August 2017

'............................................................................................................................

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
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

WshSysEnv ("WF") = "C200271"
vurl="https://"&Serverip&":58442/ecc/login.jsp"

'Import data sheet

importDataSheet()

'Import variables from sheet

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")
supprole = DataTable.Value("SupplementalRole1","Global")
username2 = DataTable.Value("User_Login","Global")
password2 = DataTable.Value("User_Password","Global")
maxpass = DataTable.Value("MaximumPasswordDuration","Global")
advancepass = DataTable.Value("AdvanceWarningofExpiration","Global")
userlicenseadvanced = DataTable.Value("Userlicenseadvancedwarning","Global")
userlicensenumberwarnings = DataTable.Value("Userlicensenumberwarning","Global")
supproleadvanced = DataTable.Value("SupplementalRoleadvancedwarning","Global")
supprolenumberwarning = DataTable.Value("SupplementalRolenumberwarning","Global")


var1 = DateAdd("d",30,date)
a = split(var1,"-")
suppexpdate = a(1)&"-"&a(0)&"-"&a(2)


'Call actions here

CreateNewUserEPS()
Wait(Iteration_Wait)
LaunchIE(vurl)
Wait(Iteration_Wait)
CertificateError()
Wait(Iteration_Wait)
ECCLogin username,password
Wait(Iteration_Wait)
AddSupplementalRole(supprole)
Wait(Iteration_Wait)
PingECC()
Wait(Iteration_Wait)
AddSupplementalRoleEPS username2,supprole,suppexpdate
Wait(Iteration_Wait)
AddLicenseUserEPS(username2)
Wait(Iteration_Wait)
passwordapplicationsettings maxpass,advancepass
Wait(Iteration_Wait)
UserLicesnseExpiration("Enabled")
Wait(Iteration_Wait)
If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If
Wait(Iteration_Wait)
UserLicenseApplicationSettings userlicensenumberwarnings,userlicenseadvanced
Wait(Iteration_Wait)
If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If
Wait(Iteration_Wait)
SupplementalRoleExpiration("Enabled")
Wait(Iteration_Wait)
If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If
Wait(Iteration_Wait)
SupplementalRoleApplicationSettings supproleadvanced,supprolenumberwarning
Wait(Iteration_Wait)
If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If
Wait(Iteration_Wait)
RunAction "Action1 [C200271_DB_Changes]", oneIteration
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username2,password2
Wait(Iteration_Wait)



If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.(st)") @@ hightlight id_;_1451390_;_script infofile_;_ZIP::ssf105.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
ElseIf JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaStaticText("Your password will expire").Check CheckPoint("Your password will expire in 2 days. Reset your password.(st)") @@ hightlight id_;_11809183_;_script infofile_;_ZIP::ssf107.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaButton("OK").Click
End If

Wait(Iteration_Wait)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.(st)") @@ hightlight id_;_1451390_;_script infofile_;_ZIP::ssf105.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
ElseIf JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaStaticText("Your password will expire").Check CheckPoint("Your password will expire in 2 days. Reset your password.(st)") @@ hightlight id_;_11809183_;_script infofile_;_ZIP::ssf107.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Authentication").JavaButton("OK").Click
End If

Wait(Iteration_Wait)

UserLicesnseExpiration("Disabled")
Wait(Iteration_Wait)
SupplementalRoleExpiration("Disabled")
Wait(Iteration_Wait)
EPSLogout()
Wait(Iteration_Wait)
EPSLogin username,password
Wait(Iteration_Wait)
RunAction "Action1 [C200271_RevertChanges]", oneIteration
Wait(Iteration_Wait)
passwordapplicationsettings "365","3"


Reporter.ReportEvent micDone,"C200271","Test case has been executed" @@ hightlight id_;_12332649_;_script infofile_;_ZIP::ssf1.xml_;_
