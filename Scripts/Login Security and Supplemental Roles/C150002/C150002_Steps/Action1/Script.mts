'...........................................................................................................................................

'Test Name : C150002_Steps

'...........................................................................................................................................

Set WshShell = CreateObject("WScript.Shell")
Set WshSysEnv = WshShell.Environment("User")

username = DataTable.Value("UserName","Global")
password = DataTable.Value("Password","Global")
username1 = WshSysEnv("UserEmpID")
password1 = DataTable.Value("User_Password","Global")

'..................................................Scenario 1...............................................................................


SupplementalRoleExpiration("Enabled")
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")
Wait(Iteration_Wait)
 @@ hightlight id_;_23140883_;_script infofile_;_ZIP::ssf7.xml_;_
 
If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If 

EPSLogin username1,password1

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaTable("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.") @@ hightlight id_;_32543891_;_script infofile_;_ZIP::ssf14.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If

'............................................................................................................................................

'..........................................................Scenario 2......................................................................

SupplementalRoleExpiration("Disabled")
Wait(Iteration_Wait)
UserLicesnseExpiration("Enabled")
Wait(Iteration_Wait)
 @@ hightlight id_;_23140883_;_script infofile_;_ZIP::ssf7.xml_;_
 
If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If 

EPSLogin username1,password1

Wait(2)
 
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.(st)") @@ hightlight id_;_5411871_;_script infofile_;_ZIP::ssf16.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If


'............................................................................................................................................

'..........................................................Scenario 3......................................................................

SupplementalRoleExpiration("Enabled")
Wait(Iteration_Wait)
UserLicesnseExpiration("Enabled")
Wait(Iteration_Wait)
 @@ hightlight id_;_23140883_;_script infofile_;_ZIP::ssf7.xml_;_
 
If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If 

EPSLogin username1,password1

Wait(2)



JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaStaticText("User privileges below").Check CheckPoint("User privileges below are either expired or about to expire. Check with your pharmacy manager or system administrator.(st)_2") @@ hightlight id_;_5075316_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("User Login Alert").JavaButton("OK").Click
End If

'.........................................................................................................................................

SupplementalRoleExpiration("Disabled")
Wait(Iteration_Wait)
UserLicesnseExpiration("Disabled")


'Login with akumar user

If JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Logout").Click
End If 

EPSLogin username,password



Reporter.ReportEvent micDone,"Steps","All steps have been executed"
