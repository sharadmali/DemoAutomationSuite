'=========================================================================================================
'Variable Declaration
'ALM Framework
'=========================================================================================================
Dim strCompName, blnExecute, blnReplaceHost, intRowCount
Dim clsDriverScript					   'clsEnvironmentVariables ' Local Variable Declaration
Dim strCurDir								'To hold the value of Current Directory
Dim strBaseDir							 'To hold the value of Base Directory cust
Dim strTempName   				  'To hold the name of file to load into runtime datatable. 
Dim strStartTime		  				'To hold the value of start time of the TestCase execution
Dim strEndTime			 				'To hold the value of end time of the TestCase execution 
Dim nTransactionCnt         		'To hold the no. of rows in the ControlFile
Dim nTemp              				         'Used as a temporary variable to hold intermediate data
Dim nLoopCnt            				  'Used as a counter
Dim flrptFile            					    'To hold the file object
Dim strPrevTSR       				   'To hold previous object repository
Dim Syntime      						    'Synt time out
Dim arrGroup()							   'To hold the values of groups from database	
Dim bResult							         'To hold the bollean value of test cases result, if test case fail it will be False, if pass then True
Dim bres										'To hold the bollean value of test cases result, if test case fail it will be False, if pass then True
Dim gstrExecutedBy,gstrRunBy,gstrHostName
Dim strQCIntegrationArray()   
Dim iSheetCnt
Dim nArrSize,nLoopCounter,gTempTCRowsCnt, strGroupNameAppIdentifier,UpLoadExcelReport
Dim strEnvVal,strApplication
Dim objApp
Dim SuiteName 
Dim strFilesystemPath

gTempTCRowsCnt=0
UpLoadExcelReport=True
gbScreenShotImageStatus=True
Set fso = CreateObject("Scripting.FileSystemObject")
Set dictApplicationURL=CreateObject("Scripting.Dictionary")
Set gDictIBCData=CreateObject("Scripting.Dictionary")
Set gDictUser=CreateObject("Scripting.Dictionary")
Set gDictPassword=CreateObject("Scripting.Dictionary")
Set gDictAdam=CreateObject("Scripting.Dictionary")
Set gDictEnvironment=CreateObject("Scripting.Dictionary")
Set dictPuttyEnv=CreateObject("Scripting.Dictionary")
SystemUtil.CloseProcessByName("EXCEL.EXE")
SystemUtil.CloseProcessByName("WINWORD.EXE")

On error resume next

' Desc: ScriptPath,  Store the complete path of Run Manager
ScriptPath = Environment("TestDir") 
DriveName = fso.GetDriveName(ScriptPath)
DriveName = DriveName & "\"
strProjectName = "Demo"
gstrProjectName = strProjectName
gstrProjectUser = Right(Environment.Value("TestName"), Len(Environment.Value("TestName")) - Len("RunManager"))
QCSuitePath="Subject\"	  
strControlPath = Left(ScriptPath,Len(ScriptPath)-10 - Len(gstrProjectUser))   
SuiteName=UCase(strProjectName) & "AutomationSuite"
gstrBaseDir = strControlPath

'Set objApp = CreateObject("QuickTest.Application")
'objApp.Folders.RemoveAll
'objApp.Folders.Add(gstrBaseDir)
'Set objApp = Nothing
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Reporting Folder structure
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
gstrEnvironmentSetupDir =  gstrBaseDir & "Libraries"
gstrLibrariesDir =  gstrBaseDir & "Libraries"
gstrObjectRepositoryDir = gstrBaseDir & "ObjectRepository"
gstrRuntimeReportsDir = gstrBaseDir & "Reports\" & gstrProjectUser & "\RuntimeReport\"
gstrScreenshotsDir = gstrBaseDir & "Reports\" & gstrProjectUser & "\Screenshots\"
gstrSummaryReportDir = gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\"
gstrTestResultLogDir = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\"
gstrActionFilesDir = gstrBaseDir & "TestArtifacts\Actions"
gstrBatchFilesDir = gstrBaseDir & "TestArtifacts\"
gstrControlFilesDir = gstrBaseDir & "TestArtifacts\"
gstrTestDataDir = gstrBaseDir & "TestArtifacts\"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Global exist wait count
gExistCount = 20

rCurTime = date & "_" & time

rCurTime = replace(rCurTime,"/","_")
rCurTime = replace(rCurTime,":","_")
rCurTime = replace(rCurTime," ","_")

Reporter.Filter = rfDisableAll
i=0
On Error Goto 0
gQCIntegration_Flag= False

'Get Environment Details
Call GetEnvironment()
Call Setup()

strCntrlPath=strControlPath
clsEnvironmentVariables.CntrlPath = strCntrlPath
 
'Initializes all the environment parameters.
clsEnvironmentVariables.Connection_String=temp_hold
clsEnvironmentVariables.ProjectName = strProjectName
clsEnvironmentVariables.MasterXLSPath =  gstrBaseDir & "TestArtifacts\GroupControlFiles" & gstrProjectUser & ".xls"
clsEnvironmentVariables.ControlPath = gstrControlFileDir
clsEnvironmentVariables.TempFilePath = Environment("SystemTempDir")
clsEnvironmentVariables.TempSummaryFilePath = Environment("ResultDir")
clsEnvironmentVariables.RunTimeReportPath = gstrRuntimeReportsDir
clsEnvironmentVariables.TestResultPath = gstrTestResultLogDir
clsEnvironmentVariables.SummaryResultPath =  gstrSummaryReportDir
gstrRunBy = Environment("UserName")

If blnQCUpdation Then
		Call clsQCIntegrationModule.ConnectToQC( gstrQCIntegrationArray(0), gstrQCIntegrationArray(1), gstrQCIntegrationArray(2),gstrQCIntegrationArray(3),gstrQCIntegrationArray(4))
End If

Call UNLOADOR()
gstrTester =  Environment.Value( "UserName" )
gstrMachine =  Environment.Value( "LocalHostName" )
gstrRunDate = Date()
gstrRunStartTime = Time() 
gstrProject=strProjectName

Call ReadGroupData(arrGroup)
nArrSize=Ubound(arrGroup)

For nLoopCounter=0 to nArrSize' To read the GroupScenario File one after another
	'Read The Control File    
	gstrAppName = arrGroup(nLoopCounter)' To read the current Scenario Group
	gstrGroupName = gstrAppName
	'ControlFile Name should be <GroupName>_Scenario.xls
	gstrControlFileName = gstrBaseDir & "TestArtifacts\GroupControlFiles" & gstrProjectUser & ".xls"       
	'Function call to load the ControlFile in RunTime datatable
	importSheetsToRunTimeDataTable gstrControlFileName,gstrGroupName,gstrGroupName
	
	nTransactionCnt =  DataTable.GetSheet(gstrGroupName).GetRowCount	
	'Looping through the contro file and executing batch files related to scenarios if execution status is "Y"  
	'Initialising Counts
	gTotalCount = 0
	gTotalPassed = 0
	gTotalFailed = 0
	gGroupExecutionTime = 0
	
	For nLoopCnt=1 To  nTransactionCnt
		Set fso = CreateObject("Scripting.FileSystemObject") 
		DataTable.GetSheet(gstrGroupName).SetCurrentRow(nLoopCnt)
		'Checking Execution status
		If  Trim ( Ucase( DataTable( "Execute",gstrGroupName ) ) ) = "Y" Then
			gTotalCount = gTotalCount + 1
			bResult = True
			gIterCombinedStatus = True
			gScreenShotPath = ""
			'gstrDataReusabilityStatus = DataTable( "DataReusabilityStatus", gstrGroupName)
			'gstrDataTestCaseID = DataTable( "DataTestCaseID", gstrGroupName)   	  
			
			If  gstrDataReusabilityStatus ="NONREUSABLE" OR gstrDataReusabilityStatus="EXECUTEDATATESTCASE" Then
					arrdata=Split(gstrDataTestCaseID,":") 
					strID=arrdata(0)
				 	strSheet=arrdata(1)
					gstrCurScenario = DataTable( "TestCaseID", gstrGroupName )
					strBrand=Right(gstrCurScenario,3)
					bresult = executedatatestcase(strID,strSheet,strBrand,UpLoadExcelReport)
			End If
			
			'Getting scenario Information and storing in global variables
			gstrCurID = DataTable( "ID", gstrGroupName )
			gstrCurScenario = DataTable( "TestCaseID", gstrGroupName )
			gstrcurScenarioDesc =  DataTable( "Description", gstrGroupName )
			gstrCurBatch = DataTable("Batch_Test_File",gstrGroupName)	
			ExcelFilePath= gstrBaseDir &"Reports\" & gstrProjectUser & "\TestResultLog\"
			ExcelFileWithPath = ExcelFilePath & "TestResults.xls"
			If (fso.FileExists(ExcelFileWithPath)) Then
						fso.deletefile(ExcelFileWithPath)
			End If
			WriteHTMLHeader clsEnvironmentVariables
			'Marking Start time of the test case execution
			strStartTime=Left( MonthName( Month( Date() ) ), 3 ) & " " & Day( Date() ) & ", " & Year( Date() ) & " " & Time()
			If  blnMultiUserExecution Then
				If UCase(gstrHostName)=UCase(gstrMachine) Then
					If  blnQCUpdation Then
								clsEnvironmentVariables.TestResultExcelFile=ExcelFileWithPath
								Set clsEnvironmentVariables.UseExcelObject =  clsQCIntegrationModule.CreateExcelFile()
					End If
					bresult = DriverScript   ' (KeywordDriver.VBS)
				Else
					UpLoadExcelReport=False
				End If
			Else
				If  blnQCUpdation Then
					clsEnvironmentVariables.TestResultExcelFile=ExcelFileWithPath
					Set clsEnvironmentVariables.UseExcelObject =  clsQCIntegrationModule.CreateExcelFile()
				End If
				
				bresult = DriverScript   ' (KeywordDriver.VBS)
						
			End If
			
			If  UpLoadExcelReport Then
				'This part of code written for summary Report
				strEndTime=Left(MonthName(Month(Date())),3) & " " & Day(Date()) & ", " & Year(Date()) & " " & Time() 'End time of the test case execution
				ExecutionTime = DateDiff("s", strStartTime, strEndTime)' Total execution time of the test case
				gGroupExecutionTime = gGroupExecutionTime + ExecutionTime 'Total execution time of the group
				WriteHTML_Verification 
				Set fso = CreateObject("Scripting.FileSystemObject")' To create .rpt for HTML reporting
				If fso.FileExists(gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".rpt") Then
					Set flrptFile = fso.GetFile(gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".rpt")
					flrptFile.Delete(true)
				End If

				Dim gstrDesc   
				If Lcase(bresult)=Lcase("Warning") Then 
					gstrDesc=""
					WriteHTMLSummaryResultLog 3
					gTotalPassed = gTotalPassed + 1
				ElseIf bresult=False Then  
					gstrDesc=""
					WriteHTMLSummaryResultLog 0	   
					gTotalFailed = gTotalFailed + 1 			
				ElseIf bresult=True Then
					gstrDesc=""
					WriteHTMLSummaryResultLog 1	 
					gTotalPassed = gTotalPassed + 1  			 
				End If  			 			
				'Desc: To handle QC integration
				If  blnQCUpdation Then
					If  UpLoadExcelReport Then
						clsQCIntegrationModule.CloseExcelFile  clsEnvironmentVariables.UseExcelObject
						Set clsEnvironmentVariables.UseExcelObject  = Nothing
						'Upload all test results into QC from excel file
						clsQCIntegrationModule.UploadTestResultsInQCFromExcel  clsEnvironmentVariables.TestSetPathQC , clsEnvironmentVariables.TestSetNameQC,clsEnvironmentVariables.TestResultExcelFile
					End If 																	  
				End If
			End If
		End If
		
		gstrStrUsername=""
		gstrtemppswd=""
		strtempusname=""
	Next
	
	'Write Summary Report    
	Write_Summary_Header clsEnvironmentVariables	
	WriteHTML_Summary_Verification	
	arrGroup(nLoopCounter) = gstrCurSummaryPath	
	SystemUtil.CloseProcessByName("EXCEL.EXE")	 
Next

For nLoopCounter=0 to nArrSize
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run """iexplore.exe"" " & arrGroup(nLoopCounter) & "",1,False
	Set objShell = Nothing

'	SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe", arrGroup(nLoopCounter)
Next
