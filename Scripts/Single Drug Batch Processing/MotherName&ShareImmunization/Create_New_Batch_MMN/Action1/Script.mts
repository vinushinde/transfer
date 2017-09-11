'..................................................................................................................................

'Test Name : Create_New_Batch_MMN

'Test Description : This test will create a new batch for single drug batch processing

'Author : Kashish Ambwani

'..................................................................................................................................

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


WF_Value = WshSysEnv ("WF")


'Import values from sheet

pat1share = DataTable.Value("PT1_ShareImmunization","Global")
pat1mothername = DataTable.Value("PT1_MotherMaidenName","Global")
eccshare = DataTable.Value("ECC_ShareImmunization","Global")

pat1share2 = DataTable.Value("PT1_ShareImmunization2","Global")
pat1mothername2 = DataTable.Value("PT1_MotherMaidenName2","Global")


pat2share = DataTable.Value("PT2_ShareImmunization","Global")
pat2mothername = DataTable.Value("PT2_MotherMaidenName","Global")

ndcdrug = DataTable.Value("NDC","Global")

'Scenario 1 and 2

'Navigate to Tools>Utilities>Single Drug Batch Processing>Create New Batch
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Create New Batch").Select @@ hightlight id_;_17891018_;_script infofile_;_ZIP::ssf2.xml_;_

'Enter all mandatory details on Drug Selection Screen

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").WaitProperty "visible",1

'Enter DAW

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").Select DataTable.Value("DAW","Global")

'Enter Rx Written Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Set DataTable.Value("Rx_Written_Date","Global") @@ hightlight id_;_989576_;_script infofile_;_ZIP::ssf4.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Type micTab

'Enter Prescribed Drug

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Drug").Set DataTable.Value("NDC","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find").Click
End If

If  JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").GetROProperty("enabled") Then
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").Click	
End If

'Enter Prescribed Quantity

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Qty.").Set DataTable.Value("Prescribed_Qty","Global") @@ hightlight id_;_31330009_;_script infofile_;_ZIP::ssf10.xml_;_

'Enter SIG Code

JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Set DataTable.Value("SIG_Code","Global") @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf8.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Type micTab

'Search and select prescriber @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf11.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescriber Last Name").Set DataTable.Value("PRESCRIBER_LAST_NAME","Global") @@ hightlight id_;_5189556_;_script infofile_;_ZIP::ssf12.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").Set DataTable.Value("PRESCRIBER_FIRST_NAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").JavaButton("Use Local Prescriber").Click
End If


'Enter Lot number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Lot Number").Set DataTable.Value("LOT_NUMBER","Global")

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click OK on Drug Validation Pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").JavaButton("OK").Click
End If

Wait(2)

'***********************************************Start of Scenario 1**********************************************************************

'Generate Debug Images

JavaWindow("Virtual Scanner 2.0 -").JavaButton("Generate Debug Images").Click @@ hightlight id_;_30831015_;_script infofile_;_ZIP::ssf18.xml_;_

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient Last Name").Set DataTable.Value("PT_LASTNAME","Global") @@ hightlight id_;_18387578_;_script infofile_;_ZIP::ssf19.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient First Name").Set DataTable.Value("PT_FIRSTNAME","Global") @@ hightlight id_;_8115685_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

Wait(3)

'Checkpoint for Mother's Maiden Name and Share Immnunization

screenmother = JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").GetROProperty("value")
screenshare = JavaWindow("Enterprise Pharmacy System").JavaEdit("Share Immunization").GetROProperty("value")

If CStr(screenmother) = CStr(pat1mothername) Then
	Reporter.ReportEvent micPass,"Mother's Maiden Name - Present on Patient","Mother's Maiden Name is auto-populaed;Mother's Maiden Name is correct;Mother's Maiden Name is "+ screenmother +""
	Else
	Reporter.ReportEvent micFail,"Mother's Maiden Name - Present on Patient","Expected Mother's Maiden Name : "+ pat1mothername +";Observed Mother's Maiden Name : "+ screenmother +""
End If

If CStr(screenshare) = CStr(pat1share) Then
	Reporter.ReportEvent micPass,"Share Immnunization - Present on patient","Share Immunization Value is correct;Share immunization is "+ screenshare +""
	Else
	Reporter.ReportEvent micFail,"Share Immnunization - Present on patient","Expected Immunization : "+ pat1share +";Observed Immunization : "+ screenshare +""
End If

'Click on Edit Billing

If JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Edit Billing").Click
End If
 @@ hightlight id_;_17263703_;_script infofile_;_ZIP::ssf82.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaList("Primary (1) Third Party").Select "CASH"

'Click on Save

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaButton("Save").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").JavaButton("Save").Click
Else
JavaWindow("Enterprise Pharmacy System").JavaDialog("Edit Billing Information").Close
End If

Wait(2)

'Make changes on Patient 1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Patient File").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Patient File").Click
End If

'Make changes to mother's maiden name

JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").WaitProperty "visible",1

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").Set DataTable.Value("PT1_MotherMaidenName2","Global")

'Click on Save

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaButton("Save").WaitProperty "disabled",1

'Click on Additional Tab

If JavaWindow("Enterprise Pharmacy System").JavaButton("Additional").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Additional").Click
End If

'Make changes on Share Immunization

JavaWindow("Enterprise Pharmacy System").JavaList("Share Immunization").WaitProperty "visible",1

Wait(3)

JavaWindow("Enterprise Pharmacy System").JavaList("Share Immunization").Select DataTable.Value("PT1_ShareImmunization2","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaButton("Save").WaitProperty "disabled",1
 @@ hightlight id_;_6980367_;_script infofile_;_ZIP::ssf133.xml_;_
If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Task").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Task").Click
End If

Wait(2)

'Checkpoint for Mother's Maiden Name and Share Immnunization

screenmotherchange = JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").GetROProperty("value")
screensharechange = JavaWindow("Enterprise Pharmacy System").JavaEdit("Share Immunization").GetROProperty("value")

If CStr(screenmotherchange) = CStr(pat1mothername2) Then
	Reporter.ReportEvent micPass,"Mother's Maiden Name - After change","Mother's Maiden Name is correct;Mother's Maiden Name is "+ screenmotherchange +""
	Else
	Reporter.ReportEvent micFail,"Mother's Maiden Name - Present on Patient","Expected Mother's Maiden Name : "+ pat1mothername2 +";Observed Mother's Maiden Name : "+ screenmotherchange +""
End If

If CStr(screensharechange) = CStr(pat1share2) Then
	Reporter.ReportEvent micPass,"Share Immnunization - After change","Share Immunization Value is correct;Share immunization is "+ screensharechange +""
	Else
	Reporter.ReportEvent micFail,"Share Immnunization - Present on patient","Expected Immunization : "+ pat1share2 +";Observed Immunization : "+ screensharechange +""
End If



'Click on Transmit

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Click
End If

'Checkpoint for Next button
	
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").WaitProperty "enabled",1
	
If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").GetROProperty("enabled") = 1 Then
	Reporter.ReportEvent micPass,"Next Button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Next Button","Disabled"
End If	

Wait(2)

'*************************************************End of Scenario 1******************************************************************8

'For Patient 2

'Generate Debug Images

JavaWindow("Virtual Scanner 2.0 -").JavaButton("Generate Debug Images").Click @@ hightlight id_;_30831015_;_script infofile_;_ZIP::ssf18.xml_;_

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient Last Name").Set DataTable.Value("PT_LASTNAME2","Global") @@ hightlight id_;_18387578_;_script infofile_;_ZIP::ssf19.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient First Name").Set DataTable.Value("PT_FIRSTNAME2","Global") @@ hightlight id_;_8115685_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

Wait(3)

'Checkpoint for patient 2 for share immunization

screenshare2 = JavaWindow("Enterprise Pharmacy System").JavaEdit("Share Immunization").GetROProperty("value")

If CStr(screenshare2) = CStr(eccshare) Then
	Reporter.ReportEvent micPass,"Share Immunization - Present on ECC","Share Immunization is correct;Share Immunization is "+ screenshare2 +""
	Else
	Reporter.ReportEvent micFail,"Share Immunization - Present on ECC","Expected immunization : "+ eccshare +";Observed immunization : "+ screenshare2 +""
End If

'Checkpoint for Transmit and Next button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micFail,"Transmit Button","Enabled"
	Else
	Reporter.ReportEvent micPass,"Transmit Button","Disabled"
End If

Wait(2)

'Make changes on the patient profile

If JavaWindow("Enterprise Pharmacy System").JavaButton("Patient File").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Patient File").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").WaitProperty "visible",1

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").Set pat2mothername

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaButton("Save").WaitProperty "disabled",1

'Click on Additional Tab button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Additional").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Additional").Click
End If

'Change Share Immunization value

JavaWindow("Enterprise Pharmacy System").JavaList("Share Immunization").WaitProperty "visible",1

Wait(2)

JavaWindow("Enterprise Pharmacy System").JavaList("Share Immunization").Select pat2share

If JavaWindow("Enterprise Pharmacy System").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Save").Click
End If

JavaWindow("Enterprise Pharmacy System").JavaButton("Save").WaitProperty "disabled",1

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Task").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Task").Click
End If

'Checkpoint for patient 2 mother's maiden name and share immunization

screenshare3 = JavaWindow("Enterprise Pharmacy System").JavaEdit("Share Immunization").GetROProperty("value")
screenmother3 = JavaWindow("Enterprise Pharmacy System").JavaEdit("Mother's Maiden Name").GetROProperty("value")

If CStr(screenshare3) = CStr(pat2share) Then
	Reporter.ReportEvent micPass,"Share Immunization - Value Changed","Share Immunization is correct;Share Immunization is "+ screenshare3 +""
	Else
	Reporter.ReportEvent micFail,"Share Immunization - Value Changed","Expected Immunization : "+ pat2share +" ; Observed Immunization : "+ screenshare3 +""
End If

If CStr(screenmother3) = CStr(pat2mothername) Then
	Reporter.ReportEvent micPass,"Mother Maiden Name - Added at runtime","Mother's Maiden Name is correct;Mother's Maiden is "+ screenmother3 +""
	Else
	Reporter.ReportEvent micFail,"Mother Maiden Name - Added at runtime","Expected Mother's Name : "+ pat2mothername +" ; Observed Mother's Name : "+ screenmother3 +""
End If

'Checkpoint for Transmit and Next button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micPass,"Transmit Button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Transmit Button","Disabled"
End If

'Click on Transmit

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Click
End If

Wait(5)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micPass,"Next Button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Next Button","Disabled"
End If

'******************************************************End of Scenario 2******************************************************************

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click on complete

If JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Click
End If

'Enter User Authentication

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If

Wait(Iteration_Wait)

'DB Validations for patient 1 and patient 2


pat1lname = CStr(DataTable.Value("PT_LASTNAME","Global"))
pat1fname = CStr(DataTable.Value("PT_FIRSTNAME","Global"))

pat2lname = CStr(DataTable.Value("PT_LASTNAME2","Global"))
pat2fname = CStr(DataTable.Value("PT_FIRSTNAME2","Global"))


'Query to fetch the highest Rx Number for patient 1
Rxnumber1 = "Select max("+ vSchema +"RX_SUMMARY.RX_NUMBER) RXNUM1 from "+ vSchema +"RX_SUMMARY,"+ vSchema +"PATIENT where "+ vSchema +"RX_SUMMARY.PATIENT_ID = "+ vSchema +"PATIENT.ID and "+ vSchema +"PATIENT.LAST_NAME = '"+ pat1lname +"' and "+ vSchema +"PATIENT.FIRST_NAME = '"+ pat1fname +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Rxnumber1)

'assigning value to rxnumber
rxnumber1 = rcEPSRecordSet.Fields("RXNUM1")
strrxnumber1 = CStr(rxnumber1)

Wait(3)

'Query to fetch the highest Rx Number for patient 2
Rxnumber2 = "Select max("+ vSchema +"RX_SUMMARY.RX_NUMBER) RXNUM2 from "+ vSchema +"RX_SUMMARY,"+ vSchema +"PATIENT where "+ vSchema +"RX_SUMMARY.PATIENT_ID = "+ vSchema +"PATIENT.ID and "+ vSchema +"PATIENT.LAST_NAME = '"+ pat2lname +"' and "+ vSchema +"PATIENT.FIRST_NAME = '"+ pat2fname +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Rxnumber2)

'assigning value to rxnumber
rxnumber2 = rcEPSRecordSet.Fields("RXNUM2")
strrxnumber2 = CStr(rxnumber2)

Wait(3)

'Query to fetch mother's maiden name for patient 1 and patient 2

Mother1 = "select ("+ vSchema +"PATIENT.MOTHER_MAIDEN_NAME) MNAME1 from "+ vSchema +"PATIENT where "+ vSchema +"PATIENT.LAST_NAME = '"+ pat1lname +"' and "+ vSchema +"PATIENT.FIRST_NAME = '"+ pat1fname +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Mother1)

'assigning value to mother's maiden name for patient 1
mname1 = rcEPSRecordSet.Fields("MNAME1")

Wait(2)

'For patient 2

Mother2 = "select ("+ vSchema +"PATIENT.MOTHER_MAIDEN_NAME) MNAME2 from "+ vSchema +"PATIENT where "+ vSchema +"PATIENT.LAST_NAME = '"+ pat2lname +"' and "+ vSchema +"PATIENT.FIRST_NAME = '"+ pat2fname +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Mother2)

'assigning value to mother's maiden name for patient 2
mname2 = rcEPSRecordSet.Fields("MNAME2")

'Share Immunization for Patient 1 and Patient 2

Share1 = "select (EPS2.RX_TX.IMM_SHARE_FLAG) SHARE1 from EPS2.RX_TX,EPS2.RX_SUMMARY where EPS2.RX_TX.RX_SUMMARY_ID = EPS2.RX_SUMMARY.ID and EPS2.RX_SUMMARY.RX_NUMBER = '"+ strrxnumber1 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Share1)

'assigning value to immunization flag for patient 1
share1 = rcEPSRecordSet.Fields("SHARE1")
strshare1 = CStr(share1)

Wait(2)

'For patient 2

Share2 = "select (EPS2.RX_TX.IMM_SHARE_FLAG) SHARE2 from EPS2.RX_TX,EPS2.RX_SUMMARY where EPS2.RX_TX.RX_SUMMARY_ID = EPS2.RX_SUMMARY.ID and EPS2.RX_SUMMARY.RX_NUMBER = '"+ strrxnumber2 +"'"

'Execute Query
Set rcEPSRecordSet =  EPS2DBObject.Execute(Share2)

'assigning value to immunization flag for patient 2
share2 = rcEPSRecordSet.Fields("SHARE2")
strshare2 = CStr(share2)

Wait(2)

If CStr(mname1) = CStr(pat1mothername2) Then
	Reporter.ReportEvent micPass,"Mother's Maiden Name  - Pre-populated","Correct"
	Else
	Reporter.ReportEvent micFail,"Mother's Maiden Name  - Pre-populated","Expected value : "+ pat1mothername2 +";Observed value : "+ mname1 +""
End If

Reporter.ReportEvent micPass,"Share immunization - Pre-populated",""+ strshare1 +""

'For patient 2

If CStr(mname2) = CStr(pat2mothername) Then
	Reporter.ReportEvent micPass,"Mother's Maiden Name  - Rumtime populated","Correct"
	Else
	Reporter.ReportEvent micFail,"Mother's Maiden Name  - Runtime populated","Expected value : "+ pat2mothername +";Observed value : "+ mname2 +""
End If

Reporter.ReportEvent micPass,"Share immunization - Changed Value",""+ strshare2 +""

Wait(Iteration_Wait)

'*******************************************************Start of Scenario 3**************************************************************

'Settings on ECC

'Navigate to Filecabinet>Defualt immunization Reporting @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Control Center").WebEdit("password")_;_script infofile_;_ZIP::ssf151.xml_;_

Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j_id60:anchor").Click @@ hightlight id_;_Browser("Enterprise Control Center").Page("Enterprise Control Center").WebElement("header-form:j id60:anchor")_;_script infofile_;_ZIP::ssf158.xml_;_

'Select state from list

Browser("Enterprise Control Center").Page("Enterprise Control Center").Link("Edit").WaitProperty "visible",1

rowvalue12 = Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebTable("Default Immunization Reporting").GetROProperty("rows")

Wait(1)

For f = 0 To rowvalue12 Step 1
	
	celdt1 = Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebTable("Default Immunization Reporting").GetCellData(f,"1")
	
	If Instr(celdt1,DataTable.Value("ECCImmunizationState","Global")) Then
	
	Set tem = Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebTable("Default Immunization Reporting").ChildItem(f,2,"Link",0)
	tem.Click
	Exit For
	End If
Next

'Deactivate all records

Browser("Enterprise Control Center").Page("EPS - Default Immunization").Link("Deactivate").WaitProperty "visible",1

If Browser("Enterprise Control Center").Page("EPS - Default Immunization").Link("Deactivate").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - Default Immunization").Link("Deactivate").Click
End If

Wait(3)

'Add new record with all checkbox checked

'Check all checkbox

Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebCheckBox("immReportForm:allCheckBox").Set "ON" @@ hightlight id_;_Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebCheckBox("immReportForm:allCheckBox")_;_script infofile_;_ZIP::ssf161.xml_;_

'Select value from immunization dropdown

Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebList("immReportForm:share").Select DataTable.Value("ECC_ShareImmunization","Global")

Wait(2)

If Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebButton("Add").Exist(15) Then
Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebButton("Add").Click
End If

Wait(2)

'Click on Home button

Browser("Enterprise Control Center").Page("EPS - Default Immunization").WebElement("WebTable").Click

'Ping ECC on EPS

PingECC()

'set all settings for single drug batch processing in application settings as No

SDBPApplicationSettings_2609 "No","No","No","No","No","No","No","No","No","No","No"

Wait(Min_Wait)

'Navigate to Tools>Utilities>Single Drug Batch Processing>Create New Batch
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Create New Batch").Select @@ hightlight id_;_17891018_;_script infofile_;_ZIP::ssf2.xml_;_

'Enter all mandatory details on Drug Selection Screen

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").WaitProperty "visible",1

'Enter DAW

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").Select DataTable.Value("DAW","Global")

'Enter Rx Written Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Set DataTable.Value("Rx_Written_Date","Global") @@ hightlight id_;_989576_;_script infofile_;_ZIP::ssf4.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Type micTab

'Enter Prescribed Drug

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Drug").Set DataTable.Value("NDC","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find").Click
End If

If  JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").GetROProperty("enabled") Then
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").Click	
End If

'Enter Prescribed Quantity

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Qty.").Set DataTable.Value("Prescribed_Qty","Global") @@ hightlight id_;_31330009_;_script infofile_;_ZIP::ssf10.xml_;_

'Enter SIG Code

JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Set DataTable.Value("SIG_Code","Global") @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf8.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Type micTab

'Search and select prescriber @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf11.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescriber Last Name").Set DataTable.Value("PRESCRIBER_LAST_NAME","Global") @@ hightlight id_;_5189556_;_script infofile_;_ZIP::ssf12.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").Set DataTable.Value("PRESCRIBER_FIRST_NAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").JavaButton("Use Local Prescriber").Click
End If


'Enter Lot number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Lot Number").Set DataTable.Value("LOT_NUMBER","Global")

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click OK on Drug Validation Pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").JavaButton("OK").Click
End If

Wait(2)

'Generate Debug Images

JavaWindow("Virtual Scanner 2.0 -").JavaButton("Generate Debug Images").Click @@ hightlight id_;_30831015_;_script infofile_;_ZIP::ssf18.xml_;_

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient Last Name").Set DataTable.Value("PT_LASTNAME3","Global") @@ hightlight id_;_18387578_;_script infofile_;_ZIP::ssf19.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient First Name").Set DataTable.Value("PT_FIRSTNAME3","Global") @@ hightlight id_;_8115685_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

Wait(2)

'Checkpoint for share immunization

screenshareall = JavaWindow("Enterprise Pharmacy System").JavaEdit("Share Immunization").GetROProperty("value")

If CStr(screenshareall) = CStr(DataTable.Value("ECC_ShareImmunization","Global")) Then
	Reporter.ReportEvent micPass,"Share immunization for all age groups","Share Immunization value is correct;Share immunization is "+ screenshareall +";All checkbox functionality on DEfault Immunization Reporting Page on ECC is working fine"
	Else
	Reporter.ReportEvent micFail,"Share immunization for all age groups","Expected value : "+ eccshare +";Observed value : "+ screenshareall +""
End If

'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

'*******************************************************End of Scenario 3****************************************************************


'**************************************************************Start of Scenario 4********************************************************

'Settings on ECC

'Navigate to Filecabinet>Defualt immunization Reporting @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Control Center").WebEdit("password")_;_script infofile_;_ZIP::ssf151.xml_;_

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebElement("header-form:j_id60:anchor").Click @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("Enterprise Control Center 2").WebElement("header-form:j id60:anchor")_;_script infofile_;_ZIP::ssf152.xml_;_

'Select state from the page for which you want to edit the Immunization ages

Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").Link("Edit").WaitProperty "visible",1

rowvalue = Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebTable("Default Immunization Reporting").GetROProperty("rows")

Wait(1)

For k = 0 To rowvalue Step 1
	
	celdt = Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebTable("Default Immunization Reporting").GetCellData(k,"1")
	
	If Instr(celdt,DataTable.Value("ECCImmunizationState","Global")) Then
	
	Set tem = Browser("Certificate Error: Navigation").Page("Enterprise Control Center_2").WebTable("Default Immunization Reporting").ChildItem(k,2,"Link",0)
	tem.Click
	Exit For
	End If
Next

'Deactivate all records

Browser("Certificate Error: Navigation").Page("EPS - Default Immunization").Link("Deactivate").WaitProperty "visible",1

If Browser("Certificate Error: Navigation").Page("EPS - Default Immunization").Link("Deactivate").Exist(15) Then
Browser("Certificate Error: Navigation").Page("EPS - Default Immunization").Link("Deactivate").Click	
End If

Wait(Min_Wait)

'Navigate to Home screen

Browser("Certificate Error: Navigation").Page("EPS - Default Immunization").WebElement("WebTable").Click @@ hightlight id_;_Browser("Certificate Error: Navigation").Page("EPS - Default Immunization").WebElement("WebTable")_;_script infofile_;_ZIP::ssf155.xml_;_

'Close IE browser

CloseIECertificate()

'Ping ECC

PingECC() @@ hightlight id_;_8937615_;_script infofile_;_ZIP::ssf157.xml_;_

'Create New Batch


'Navigate to Tools>Utilities>Single Drug Batch Processing>Create New Batch
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Create New Batch").Select @@ hightlight id_;_17891018_;_script infofile_;_ZIP::ssf2.xml_;_

'Enter all mandatory details on Drug Selection Screen

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").WaitProperty "visible",1

'Enter DAW

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").Select DataTable.Value("DAW","Global")

'Enter Rx Written Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Set DataTable.Value("Rx_Written_Date","Global") @@ hightlight id_;_989576_;_script infofile_;_ZIP::ssf4.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Type micTab

'Enter Prescribed Drug

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Drug").Set DataTable.Value("NDC","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find").Click
End If

If  JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").GetROProperty("enabled") Then
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").Click	
End If

'Enter Prescribed Quantity

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Qty.").Set DataTable.Value("Prescribed_Qty","Global") @@ hightlight id_;_31330009_;_script infofile_;_ZIP::ssf10.xml_;_

'Enter SIG Code

JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Set DataTable.Value("SIG_Code","Global") @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf8.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Type micTab

'Search and select prescriber @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf11.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescriber Last Name").Set DataTable.Value("PRESCRIBER_LAST_NAME","Global") @@ hightlight id_;_5189556_;_script infofile_;_ZIP::ssf12.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").Set DataTable.Value("PRESCRIBER_FIRST_NAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").JavaButton("Use Local Prescriber").Click
End If


'Enter Lot number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Lot Number").Set DataTable.Value("LOT_NUMBER","Global")

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click OK on Drug Validation Pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").JavaButton("OK").Click
End If

Wait(2)

'Generate Debug Images

JavaWindow("Virtual Scanner 2.0 -").JavaButton("Generate Debug Images").Click @@ hightlight id_;_30831015_;_script infofile_;_ZIP::ssf18.xml_;_

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient Last Name").Set DataTable.Value("PT_LASTNAME3","Global") @@ hightlight id_;_18387578_;_script infofile_;_ZIP::ssf19.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient First Name").Set DataTable.Value("PT_FIRSTNAME3","Global") @@ hightlight id_;_8115685_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

'Checkpoint for Transmit and Next Button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micPass,"Transmit Button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Transmit Button","Disabled"
End If

'Click on Transmit

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").Click
End If

Wait(5)

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").GetROProperty("enabled")=1 Then
	Reporter.ReportEvent micPass,"Next Button","Enabled"
	Else
	Reporter.ReportEvent micFail,"Next Button","Disabled"
End If

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click on complete

If JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Complete").Click
End If

'Enter User Authentication

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("User Name").Set DataTable.Value("USERNAME_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaEdit("Password").Set DataTable.Value("PASSWORD_1","Global")
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Enterprise Pharmacy System").JavaButton("OK").Click
End If


'*************************************************************End of Scenario 4***********************************************************

Wait(Min_Wait)

'****************************************************************Start of Scenario 5******************************************************

'set all settings for single drug batch processing in application settings as No

SDBPApplicationSettings_2609 "No","No","No","No","No","No","No","No","No","No","Yes"

Wait(Min_Wait)

DrugShareImmunization ndcdrug,"OFF"

Wait(Iteration_Wait)

'Navigate to Tools>Utilities>Single Drug Batch Processing>Create New Batch
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaMenu("Tools").JavaMenu("Utilities").JavaMenu("Single Drug Batch Processing").JavaMenu("Create New Batch").Select @@ hightlight id_;_17891018_;_script infofile_;_ZIP::ssf2.xml_;_

'Enter all mandatory details on Drug Selection Screen

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").WaitProperty "visible",1

'Enter DAW

JavaWindow("Enterprise Pharmacy System").JavaList("DAW").Select DataTable.Value("DAW","Global")

'Enter Rx Written Date

JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Set DataTable.Value("Rx_Written_Date","Global") @@ hightlight id_;_989576_;_script infofile_;_ZIP::ssf4.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("Rx Written").Type micTab

'Enter Prescribed Drug

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Drug").Set DataTable.Value("NDC","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find").Click
End If

If  JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").GetROProperty("enabled") Then
	JavaWindow("Enterprise Pharmacy System").JavaDialog("Packaged Drug Search").JavaButton("Select").Click	
End If

'Enter Prescribed Quantity

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescribed Qty.").Set DataTable.Value("Prescribed_Qty","Global") @@ hightlight id_;_31330009_;_script infofile_;_ZIP::ssf10.xml_;_

'Enter SIG Code

JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Set DataTable.Value("SIG_Code","Global") @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf8.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaEdit("SIG Code").Type micTab

'Search and select prescriber @@ hightlight id_;_27672439_;_script infofile_;_ZIP::ssf11.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Prescriber Last Name").Set DataTable.Value("PRESCRIBER_LAST_NAME","Global") @@ hightlight id_;_5189556_;_script infofile_;_ZIP::ssf12.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("First Name").Set DataTable.Value("PRESCRIBER_FIRST_NAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_2").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Select").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").JavaButton("Use Local Prescriber").Click
End If


'Enter Lot number

JavaWindow("Enterprise Pharmacy System").JavaEdit("Lot Number").Set DataTable.Value("LOT_NUMBER","Global")

'Click on Next

If JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Next").Click
End If

'Click OK on Drug Validation Pop-up

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Batch Processing - Drug").JavaButton("OK").Click
End If

Wait(2)

'Generate Debug Images

JavaWindow("Virtual Scanner 2.0 -").JavaButton("Generate Debug Images").Click @@ hightlight id_;_30831015_;_script infofile_;_ZIP::ssf18.xml_;_

'Search and select patient

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient Last Name").Set DataTable.Value("PT_LASTNAME2","Global") @@ hightlight id_;_18387578_;_script infofile_;_ZIP::ssf19.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaEdit("Patient First Name").Set DataTable.Value("PT_FIRSTNAME2","Global") @@ hightlight id_;_8115685_;_script infofile_;_ZIP::ssf20.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("find_3").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").GetROProperty("enabled") Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Patient Search").JavaButton("Select").Click
End If

Wait(3)

If JavaWindow("Enterprise Pharmacy System").JavaEdit("Share Immunization").Exist Then
	Reporter.ReportEvent micFail,"Share Immunization Field - Drug is not immunization drug","Field is visible"
	Else
	Reporter.ReportEvent micPass,"Share Immunization Field - Drug is not immunization drug","Field is not visible"
End If

Wait(2)

'Checkpoint for transmit button

If JavaWindow("Enterprise Pharmacy System").JavaButton("Transmit").GetROProperty("enabled")=1 Then
Reporter.ReportEvent micPass,"Transmit Button : Share immunization field is not visible;Share immunization field is mandatory","Enabled"
Else
Reporter.ReportEvent micFail,"Transmit Button : Share immunization field is not visible;Share immunization field is mandatory","Disabled"
End If


'Click on Back to Home

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(10) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

''****************************************************End of Scenario 5*******************************************************************
'
Reporter.ReportEvent micDone,"Create New Batch","New Batch for single drug batch processing has been created"
