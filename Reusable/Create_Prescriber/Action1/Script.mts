'.................................................................................................................................

'Test Name : Create_Prescriber

'Test Description : This test will create a new prescriber if it does not already exist

'Author : Kashish Ambwani

'.................................................................................................................................

'Navigate to Filecabinet>Prescriber>Information

JavaWindow("Enterprise Pharmacy System").JavaMenu("Filecabinet").JavaMenu("Prescriber").JavaMenu("Information").Select @@ hightlight id_;_28586233_;_script infofile_;_ZIP::ssf1.xml_;_

'Search for prescriber

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaEdit("Last Name").WaitProperty "visible",1

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaEdit("Last Name").Set DataTable.Value("PRESCRIBER_LAST_NAME","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaEdit("First Name").Set DataTable.Value("PRESCRIBER_FIRST_NAME","Global")

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("find").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("find").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaStaticText("No matching prescribers").Exist(15) Then

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Add New Prescriber").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaButton("Add New Prescriber").Click
End If


JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaEdit("Address").Set DataTable.Value("PRESCRIBER_ADDRESS","Global")
 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf8.xml_;_
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaEdit("City").Set DataTable.Value("PRESCRIBER_CITY","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaList("State").Select DataTable.Value("PRESCRIBER_STATE","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaEdit("ZIP Code").Set DataTable.Value("PRESCRIBER_ZIPCODE","Global")

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaEdit("Primary Office Number").Set DataTable.Value("PRESCRIBER_PHONE","Global") @@ hightlight id_;_9852337_;_script infofile_;_ZIP::ssf13.xml_;_

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaEdit("Primary Office Number").Type micTab

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaButton("Save").GetROProperty("enabled")=1 Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Add New Prescriber").JavaButton("Save").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").JavaDialog("Not found in Central Prescriber").JavaButton("Add Local Prescriber").Click
End If

If JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Exist(15) Then
JavaWindow("Enterprise Pharmacy System").JavaButton("Back to Home").Click
End If

Else

JavaWindow("Enterprise Pharmacy System").JavaDialog("Prescriber Search").Close

End If @@ hightlight id_;_30512343_;_script infofile_;_ZIP::ssf16.xml_;_

Reporter.ReportEvent micDone,"Create New Prescriber","New prescriber has been created"
