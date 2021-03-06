﻿Class EnvironmentVariables
	
	'Declares all the private variables.
	Private strControlPath, strMasterXlsPath, strURLPath, strRunTimeReportPath, strTestResultPath, strSummaryResultPath 
	Private strTempFilePath, strTempSummaryFilePath, strBrowserName, strScenarioName, iErrNum, strProjectName, strAppType
	Private strRunBy, strTestCycle, strBuild, strScenarioPath(),blnDateFormate
	Private strToAdr, blnMail, strSubject, strMsg,strBrowserTitle
	Private strstepStatus, ConnectionString, TestCase_Status,strIterationCount
	Private Scenario_Status,strContrlPath,strScreenShotPath,strCompleteTestStepCount,strFunctionalArea
        Private App_date,strEnvironment,strEnvSheetPos,strTestCaseName,strScriptPath,strCurrentTestStepCount
	Private nIDIndex
	
	
	'Desc: To handle QC integration
	'Author: Sharad Mali
	Private objExcel, strExcelFile,objUseExcelObject
	Private strExcelFileName,strServerNameQC,strUserNameQC,strPasswordQC,strDomainQC,strProjectQC,strTestSetPathQC,strTestSetNameQC,strTestCaseNameFromQC,strTestStepCount,strTestStepAction,strActionLabelName,blnQCUpdation

	
	'Function defination to Set and Retrieve the values in the private variables.

	Public Property Let CurrentTestStepCount(StepCount)
		strCurrentTestStepCount = StepCount
	End Property
	
	Public Property Get CurrentTestStepCount()
		CurrentTestStepCount = strCurrentTestStepCount
	End Property

	Public Property Let CompleteTestStepCount(StepCount)
		strCompleteTestStepCount = StepCount
	End Property
	
	Public Property Get CompleteTestStepCount()
		CompleteTestStepCount = strCompleteTestStepCount
	End Property
	
	Public Property Let Iteration_Per_Cycle(IterationCount)
		strIterationCount = IterationCount
	End Property
	
	Public Property Get Iteration_Per_Cycle()
		Iteration_Per_Cycle = strIterationCount
	End Property

	Public Property Let ScriptPath(Script_Path)
		strScriptPath = Script_Path
	End Property
	
	Public Property Get ScriptPath()
		ScriptPath = strScriptPath
	End Property

	Public Property Let TestCaseName(TestCase)
		strTestCaseName = TestCase
	End Property
	
	Public Property Get TestCaseName()
		TestCaseName = strTestCaseName
	End Property 

	Public Property Let ScreenShotPath(SnapshotPath)
		strScreenShotPath = SnapshotPath
	End Property
	
	Public Property Get ScreenShotPath()
		ScreenShotPath = strScreenShotPath
	End Property
	
	Public Property Let CntrlPath(strCntrlPath)
		strContrlPath = strCntrlPath
	End Property
	
	Public Property Get CntrlPath()
		CntrlPath = strContrlPath
	End Property

        Public Property Let Msg(strData)
		strMsg = strData
	End Property
	
	Public Property Get Msg()
		Msg = strMsg
	End Property

	Public Property Let ToAdr(strData)
		strToAdr = strData
	End Property
	
	Public Property Get ToAdr()
		ToAdr = strToAdr
	End Property

	Public Property Let Subject(strData)
		strSubject = strData
	End Property
	
	Public Property Get Subject()
		Subject = strSubject
	End Property

	Public Property Let Mail(strData)
		blnMail = strData
	End Property
	
	Public Property Get Mail()
		Mail = blnMail
	End Property
	
	Public Property Let DateFormate(strData)
		blnDateFormate = strData
	End Property
	
	Public Property Get DateFormate()
		DateFormate = blnDateFormate
	End Property

	Public Function SetScenarioPath(ByVal strData, ByVal intInd)
		ReDim Preserve strScenarioPath(intInd)
		strScenarioPath(intInd) = strData
	End Function
	
	Public Function GetScenarioName(ByVal intInd)
		GetScenarioName = strScenarioPath(intInd)
	End Function

	Public Property Let Build(strData)
		strBuild = strData
	End Property
	
	Public Property Get Build()
		Build = strBuild
	End Property

	Public Property Let TestCycle(strData)
		strTestCycle = strData
	End Property
	
	Public Property Get TestCycle()
		TestCycle = strTestCycle
	End Property

	Public Property Let RunBy(strData)
		strRunBy = strData
	End Property
	
	Public Property Get RunBy()
		RunBy = strRunBy
	End Property

	Public Property Let Connection_String(strData)
		ConnectionString = strData
	End Property
	
	Public Property Get Connection_String()
		Connection_String = ConnectionString
	End Property
	
	Public Property Let ProjectName(strData)
		strProjectName = strData
	End Property
	
	Public Property Get ProjectName()
		ProjectName = strProjectName
	End Property

	Public Property Let AppType(strData)
		strAppType = strData
	End Property
	
	Public Property Get AppType()
		AppType = strAppType
	End Property
	
	Public Property Let ControlPath(strData)
		strControlPath = strData
	End Property
	
	Public Property Get ControlPath()
		ControlPath = strControlPath
	End Property

	Public Property Let MasterXLSPath(strData)
		strMasterXlsPath = strData
	End Property
	
	Public Property Get MasterXLSPath()
		MasterXLSPath = strMasterXlsPath
	End Property

	Public Property Let URLPath(strData)
		strURLPath = strData
	End Property
	
	Public Property Get URLPath()
		URLPath = strURLPath
	End Property

	Public Property Let Environment(Environment_Data)
		strEnvironment = Environment_Data
	End Property
	
	Public Property Get Environment()
		Environment = strEnvironment
	End Property

	Public Property Let EnvSheetPos(EnvPos)
		strEnvSheetPos = EnvPos
	End Property
	
	Public Property Get EnvSheetPos()
		EnvSheetPos = strEnvSheetPos
	End Property

	Public Property Let RunTimeReportPath(strData)
		strRunTimeReportPath = strData
	End Property
	
	Public Property Get RunTimeReportPath()
		RunTimeReportPath = strRunTimeReportPath
	End Property

	Public Property Let TestResultPath(strData)
		strTestResultPath = strData
	End Property
	
	Public Property Get TestResultPath()
		TestResultPath = strTestResultPath
	End Property
		
	Public Property Let SummaryResultPath(strData)
		strSummaryResultPath = strData
	End Property
	
	Public Property Get SummaryResultPath()
		SummaryResultPath = strSummaryResultPath
	End Property

	Public Property Let TempSummaryFilePath(strData)
		strTempSummaryFilePath = strData
	End Property
	
	Public Property Get TempSummaryFilePath()
		TempSummaryFilePath = strTempSummaryFilePath
	End Property

	Public Property Let TempFilePath(strData)
		strTempFilePath = strData
	End Property
	
	Public Property Get TempFilePath()
		TempFilePath = strTempFilePath
	End Property

	Public Property Let BrowserName(strData)
		strBrowserName = strData
	End Property
	
	Public Property Get BrowserName()
		BrowserName = strBrowserName
	End Property

	Public Property Let ScenarioName(strData)
		strScenarioName = strData
	End Property
	
	Public Property Get ScenarioName()
		ScenarioName = strScenarioName
	End Property

	Public Property Let ErrNum(strData)
		iErrNum = strData
	End Property
	
	Public Property Get ErrNum()
		ErrNum = iErrNum
	End Property
       
        Public Property Let stepStatus(strData)
		strstepStatus = strData
	End Property
		
	Public Property Get stepStatus()
		stepStatus = strstepStatus
	End Property
	                  
        Public Property Let BrowserTitle(strData)
		strBrowserTitle = strData
	End Property
		
	Public Property Get BrowserTitle()
		BrowserTitle = strBrowserTitle
	End Property

	Public Property Let TestCaseStatus(RunStatus)
		TestCase_Status=CBool(RunStatus)
	End Property

	Public Property Get TestCaseStatus()
		TestCaseStatus=TestCase_Status
	End Property

	'Desc: To handle QC integration
	'Author: Sharad Mali
	Public Property Set UseExcelObject(objExcel)
		Set objUseExcelObject=objExcel
	End Property

	Public Property Get UseExcelObject()
		Set UseExcelObject=objUseExcelObject
	End Property

	Public Property Let TestResultExcelFile(strData)
		strExcelFileName=strData
	End Property

	Public Property Get TestResultExcelFile()
		TestResultExcelFile=strExcelFileName
	End Property

	Public Property Let ServerNameQC(strData)
		strServerNameQC=strData
	End Property

	Public Property Get ServerNameQC()
		ServerNameQC=strServerNameQC
	End Property

	Public Property Let UserNameQC(strData)
		strUserNameQC=strData
	End Property

	Public Property Get UserNameQC()
		UserNameQC=strUserNameQC
	End Property	
	
	Public Property Let PasswordQC(strData)
		strPasswordQC=strData
	End Property

	Public Property Get PasswordQC()
		PasswordQC=strPasswordQC
	End Property
	
	Public Property Let DomainQC(strData)
		strDomainQC=strData
	End Property

	Public Property Get DomainQC()
		DomainQC=strDomainQC
	End Property

	Public Property Let ProjectQC(strData)
		strProjectQC=strData
	End Property

	Public Property Get ProjectQC()
		ProjectQC=strProjectQC
	End Property

	Public Property Let FolderNameQC(strData)
		strFolderNameQC=strData
	End Property

	Public Property Get FolderNameQC()
		FolderNameQC=strFolderNameQC
	End Property

	Public Property Let TestSetPathQC(strData)
		strTestSetPathQC=strData
	End Property

	Public Property Get TestSetPathQC()
		TestSetPathQC=strTestSetPathQC
	End Property

	Public Property Let TestSetNameQC(strData)
		strTestSetNameQC=strData
	End Property

	Public Property Get TestSetNameQC()
		TestSetNameQC=strTestSetNameQC
	End Property
	
	Public Property Let TestCaseNameFromQC(strData)
		strTestCaseNameFromQC=strData
	End Property
	Public Property Get TestCaseNameFromQC()
		TestCaseNameFromQC=strTestCaseNameFromQC
	End Property
	
	Public Property Let TestStepCount(strData)
		strTestStepCount=strData
	End Property

	Public Property Get TestStepCount()
		TestStepCount=strTestStepCount
	End Property

	Public Property Let TestStepAction(strData)
		strTestStepAction=strData
	End Property

	Public Property Get TestStepAction()
		TestStepAction=strTestStepAction
	End Property

	Public Property Let ActionLabelName(strData)
		strActionLabelName=strData
	End Property

	Public Property Get ActionLabelName()
		ActionLabelName=strActionLabelName
	End Property

	Public Property Let QCUpdation(Data)
		blnQCUpdation=CBool(Data)
	End Property

	Public Property Get QCUpdation()
		QCUpdation=blnQCUpdation
	End Property

	Public Property Let IDIndex(StepCount)
		nIDIndex = StepCount
	End Property
	
	Public Property Get IDIndex()
		IDIndex = nIDIndex
	End Property

	  
End Class
