'......................................................................................................................................

'Test Name : GroupIncludeExclude_AT_RF

'Test Description : This test case includes 4 defects which are mentioned below : 
					'1. EPSCO - 863 : Group Include/Exclude - Default Focus is not provide on page for keyboard tabbing
					'2. EPSCO - 864 : Group Include/Exclude - Page throws Exception on keyboard Tabbing
					'3. EPSCO - 1530 : Ins Plan Group Include/Exclude - Reverse Tabbing
					'4. EPSCO - 1531 : Group Include/Exclude - No error message for Group ID Max Length

'Author : Kashish Ambwani

'Date modified : 20th June 2017

'......................................................................................................................................

DataTable.ImportSheet "E:\Sprint\Scripts\Costco\Tab Order\Group Include-Exclude\GroupIncludeExclude.xls",1,"Global"


'Navigate to Filecabinet>Insurance Plan>Information

JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Insurance Plan").JavaMenu("Information").Select @@ hightlight id_;_5727622_;_script infofile_;_ZIP::ssf1.xml_;_

'Search and select Third Party

JavaWindow("Enterprise Pharmacy System").JavaDialog("Insurance Plan Search").JavaEdit("Carrier ID").Set DataTable.Value("ThirdParty","Global")

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Insurance Plan Search").JavaButton("find").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Insurance Plan Search").JavaButton("find").Click
End If

Wait(2)

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Insurance Plan Search").JavaButton("Select").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Insurance Plan Search").JavaButton("Select").Click
End If

Wait(3)

'Navigate to Group Include/Exclude through Keyboard inputs
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf5.xml_;_
'Keyboard input for Alt+F9
 
JavaWindow("Enterprise Pharmacy System").Type micAlt+micF9

Wait(1)

JavaWindow("Enterprise Pharmacy System").Type micCtrl+micAlt+"G"

JavaWindow("Enterprise Pharmacy System").JavaEdit("Group ID").WaitProperty "visible",1

Wait(1)

If JavaWindow("Enterprise Pharmacy System").GetROProperty("focused") = 1  Then
	Reporter.ReportEvent micPass,"Focus","Focus is present on screen"
	Else
	Reporter.ReportEvent micFail,"Focus","Focus is not present on screen"
End If

JavaWindow("Enterprise Pharmacy System").Type micTab

If JavaWindow("Enterprise Pharmacy System").GetROProperty("focused") = 1  Then
	Reporter.ReportEvent micPass,"Focus","Focus is present on screen"
	Else
	Reporter.ReportEvent micFail,"Focus","Focus is not present on screen"
End If

JavaWindow("Enterprise Pharmacy System").Type micShiftDwn+micTab


If JavaWindow("Enterprise Pharmacy System").GetROProperty("focused") = 1  Then
	Reporter.ReportEvent micPass,"Focus","Focus is present on screen"
	Else
	Reporter.ReportEvent micFail,"Focus","Focus is not present on screen"
End If


JavaWindow("Enterprise Pharmacy System").JavaEdit("Group ID").Set DataTable.Value("GroupId","Global")

If JavaWindow("Enterprise Pharmacy System").JavaButton("Apply").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Apply").Click
End If

rowval = JavaWindow("Enterprise Pharmacy System").JavaTable("Group Include/Excludes").GetROProperty("rows")

For i = 0 To rowval-1 Step 1
	
	groupid = JavaWindow("Enterprise Pharmacy System").JavaTable("Group Include/Excludes").GetCellData(i,"Group ID")
	
	If groupid = DataTable.Value("GroupId","Global") Then
		Reporter.ReportEvent micPass,"Group ID","Group ID has been added"
		Exit For
		
	End If
	
Next

'Click on Back to Home @@ hightlight id_;_5867086_;_script infofile_;_ZIP::ssf11.xml_;_

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Reporter.ReportEvent micDone,"Group Include/Exclude","Completed"
