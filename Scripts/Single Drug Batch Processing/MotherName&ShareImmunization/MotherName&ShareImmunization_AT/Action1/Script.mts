'.........................................................................................................................................

'Test Name : MotherName&ShareImmunization_AT

'Test Description : This test will test the functionality of Mother's Maiden Name and Share Immunization in Single Drug Batch Processing

'Author : Kahish Ambwani

'Date Modified : 8 August 2017

'.........................................................................................................................................

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
	WF_mode = WshSysEnv("vMODE")
	
' Set the connection string and open the connection and Open the connection.
EPS_SprintDBConnection EPS2DBObject,serverip,dbuser,dbpassword

WshSysEnv ("WF") = "MotherShare"
WshSysEnv("rownum") = 1

'Import data sheet

importDataSheet()

'Import variables from sheet

pat1lname = DataTable.Value("PT_LASTNAME","Global")
pat1fname = DataTable.Value("PT_FIRSTNAME","Global")
pat1share = DataTable.Value("PT1_ShareImmunization","Global")
pat1mothername = DataTable.Value("PT1_MotherMaidenName","Global")

pat2lname = DataTable.Value("PT_LASTNAME2","Global")
pat2fname = DataTable.Value("PT_FIRSTNAME2","Global")


'Call actions here

RunAction "Action1 [Test_Data_MMN]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Preconditions_MMN]", oneIteration
Wait(Iteration_Wait)
RunAction "Action1 [Preconditions_ECC_MMN]", oneIteration
Wait(Iteration_Wait)
PatShareImmunization pat1lname,pat1fname,pat1share
Wait(Iteration_Wait)
PatShareImmunization pat2lname,pat2fname,"#0"
Wait(Iteration_Wait)
PatMotherMaidenName pat1lname,pat1fname,pat1mothername
Wait(Iteration_Wait)
RunAction "Action1 [Create_New_Batch_MMN]", oneIteration
Wait(Iteration_Wait)
PatMotherMaidenName pat2lname,pat2fname,""
Wait(Iteration_Wait)
RunAction "Action1 [SDBP_Open_Orders_MMN]", oneIteration


Reporter.ReportEvent micDone,"Mother Maiden Name and Share Immunization","This functionality has been tested"
