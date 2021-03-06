﻿'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'Library File	Name		:Keyword Driver
'Author								:Sharad Mali
'Created date					:
'Description					:It	has	all	the	function calls for various Keywords	that can be	used in	the	scripts.
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------

Option Explicit

'	Declaration	of public	variables
Public objDictTC				'Stores	a	row	of Scenario	data from	TestCase table
Public dataID					'Stores	the	data ids

'	Declaration	of local variables
Dim	dictData
'Dim	objIteration				'	To get the iterations	for	#	of of	test data	rows
Public arrData					'Stores	data for multiple	rows fetched from	database table
Dim	nScCnt						'	Number of	test data	rows in	TestCase Table
Dim	nLoopIter
Dim	nTestcaseRow
Dim	nCurrentDataRow
Dim	nCurrentDataIndex
Dim	nDataIndex
Dim	strTestData
Public nLoopDataCount

'	This function	is the heart of	the	framework.
'-------------------------------------------------------------------------------------------------
'Function	Name			:DriverScript
'Input Parameter		:None
'Description				:This	function calls the respective	functions	for	all	the	keywords used	in Test	Cases
'Calls							:All the functions listed	in Utilty.vbs
'Return	Value				:None
'-------------------------------------------------------------------------------------------------

Public Function	DriverScript

	Dim	strTempName
	Dim	nRetVal,cntpop
	nRetVal=0
	Dim	nTempLoopCnt,	nTotSheetCnt,nTempRowCnt,nTempTestCaseCnt	,nTempArrSize,nTempTCRowsCnt, nTempTCCnt
	Dim	strKeyword,	strObjName,	strParam,	strRemarks,strTestCaseName,strLabel
	Dim 	strOtherInput, SplitData, strDatabaseRef, strSplitVal, conDBRef, rsDBRef,strDBRefQuery1,strDBRefQuery2 
	Dim	strObjTableID,strTempTestCaseNames, arrParam, intCnt, strParamPart, intPos1,intPos2,strParamVal, intIterateCount, nTemp
	Dim bResult, fso
	gintCaseID=0
	
	Call readBatchFile

	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".log") Then
		fso.DeleteFile(gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".log")
	End If

	If fso.FileExists(gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".rpt") Then
		fso.DeleteFile(gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".rpt")
	End If

	Set fso = Nothing

	'Word Document:-
    gstrScreenshotBookmarkNum=1
    gstrScreenshotEND_OF_DOC=6
    gstrVariableDefine="SET"

	Set gobjWord = CreateObject( "Word.Application" )
	Set gobjDoc = gobjWord.Documents.Add()
	gobjWord.visible=False
	
	present_time=now
	present_time=Replace(present_time,"/","-")
	present_time=Replace(present_time," ","_")
	present_time=Replace(present_time,":","_")
	gstrDocName=gstrCurScenario&present_time
	gstrScreenshotDocPath=gstrScreenshotsDir&gstrDocName
	scenariodesc="Test Case Name:  "&gstrcurScenarioDesc&vbcrlf&vbcrlf
	gobjWord.Selection.TypeText(scenariodesc)
	gobjDoc.SaveAs gstrScreenshotDocPath ,0
	nTempArrSize = UBound(garrTestCaseNames) ' This will give total number of sheets in batch file

	nCurrentDataRow=0
	nCurrentDataIndex=0			
	
	'****	Iteration	Block	-	Start

	Set	objDictTC	=CreateObject("Scripting.Dictionary")
'	Set	objIteration=	new	clsIteration
	
	objIteration.getTestCaseData
	'Get the number of test data rows in the table for a scenario
	
	nScCnt = Ubound(objIteration.arrData,1)
	
	'****	Iteration Block	- End
	bResult = True
	For nLoopIter =	0 to nScCnt		'	Loop For individual	Records
		Set objDictTC = getCurrentTestCaseData	(objIteration.arrData,objIteration.arrColNames,nLoopIter)	' This gets the	Current	test data
		For nLoopDataCount = 0 to gDataCount 
			gStrTestStep=0
			For nTempTestCaseCnt =0 to nTempArrSize
		
				gstrActionName = CStr(garrTestCaseNames(nTempTestCaseCnt,0))
				gstrActionDataSet = CStr(garrTestCaseNames(nTempTestCaseCnt,1))
						
				intIterateCount = 0
		
				On Error Resume Next
				objErr.Clear
                gstrDesc =  "------------------------------------<font color=Blue>" & gstrActionName & "</font>------------------------------------" 
				WriteHTMLResultLog gstrDesc, 5
				gbReporting = True

				If gstrExecuteMethod = "QTPAction" Then
							RunAction garrTestCaseNames(nTempTestCaseCnt, 0), oneIteration
				ElseIf gstrExecuteMethod = "VBS" Then
							Execute "Call " & garrTestCaseNames(nTempTestCaseCnt, 0)
				End If
				
				If objErr.Number = 11 Then
					bResult=False
				
					If nLoopDataCount = gDataCount Then
						Exit Function
					Else
						Exit For
					End If
					
				End If
		
			Next			
			If nLoopDataCount <> gDataCount Then
				gstrDesc="<br><b><font color=Blue><i>&radic;</i> ***************ENVIRONMENT CHANGE***************</font></b>"
				WriteHTMLResultLog gstrDesc, 5
			End IF
		Next
	Next

	Call disconnectDB()

	Set objOffers=Nothing

'	Call CloseAllBrowsers("Check")
	DriverScript = bResult
	
	
End Function

'------------------------------------------------------------------------------------------------------------
'Function Name 		:readBatchFile
'Input Parameter	:None
'Description		:This function reads the batch file and loads the testcase file in datatable
'Calls			:importSheetsToRunTimeDataTable
'Return	Value		:None
'------------------------------------------------------------------------------------------------------------
Public Function readBatchFile()
   	Dim nTempColCount, conn, rs, fso, strQuery, nLoopCount, strTemp, arrTemp
  	Dim nTempIndex, arrTempAction, arrTempDataSet
	
	set conn = CreateObject("ADODB.Connection")
	set rs = CreateObject("ADODB.Recordset")
	Set fso = CreateObject("Scripting.FIleSystemObject")
	'	On Error Resume Next
	nTempColCount = 0
	If fso.FileExists(gstrControlFileName) Then
		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrBaseDir & "TestArtifacts\BatchSheet.xls" & ";Excel 8.0;HDR=Yes;"  
		strQuery = "Select * from [BatchSheet$] where BatchTestFile='" & gstrCurBatch & "'"
		rs.open strQuery,conn,1,1
		nTempColCount = rs.Fields.Count
		If rs.RecordCount > 0 Then
			nTempIndex = 0
			For nLoopCount = 1 to nTempColCount - 1
					strTemp = Trim(rs.Fields(nLoopCount).Value)
					
					If  strTemp <> "" Then
							If InStr(strTemp, ",") = 0 Then
									strTemp = strTemp & ",A"
							End If
							
							arrTemp = Split(strTemp, ",")
		
							If IsArray(arrTempAction) Then
									ReDim Preserve arrTempAction(nTempIndex)
									ReDim Preserve arrTempDataSet(nTempIndex)
							Else
									ReDim arrTempAction(nTempIndex)
									ReDim arrTempDataSet(nTempIndex)
							End If
							
							arrTempAction(nTempIndex)= Trim(arrTemp(0))
							arrTempDataSet(nTempIndex) = Trim(arrTemp(1))
							nTempIndex = nTempIndex + 1
					End If
				
			Next
			ReDim garrTestCaseNames(nTempIndex-1,1)
			For nLoopCount = 0 to nTempIndex - 1
					garrTestCaseNames(nLoopCount,0) = arrTempAction(nloopCount)
					garrTestCaseNames(nLoopCount,1) = arrTempDataSet(nloopCount)
			Next

			Erase arrTempAction
			Erase arrTempDataSet
			
		Else
			clsEnvironmentVariables.ErrNum = "#"
			WriteHTMLErrorLog clsEnvironmentVariables, "Batch File does not exist : " & gstrCurBatch , "", "", ""
		End If
		rs.Close
		conn.Close
	End If
	On Error Goto 0
	Set rs = Nothing
	Set conn = Nothing
								
End Function
'------------------------------------------------------------------------------------------------------------

Public Function getData(strParam)
						
'   	Get the Row Ids from the TestCase table. This will get all the test data required for execution of the test cases
						dataID = getDataRowIDS(strParam,objDictTC)
						gDataID = dataID
'						MsgBox dataID
						If Instr(dataID,",") <> 0 Then
							SplitData = Split(dataID,",")
							dataID = SplitData(nLoopDataCount)
						End If						
						'Get the data from the table mentioned in strParam and store in arrData
						If isArray(arrData) Then
							Erase arrData
						End If
						
						Set dictData = getDataArray(strParam,dataID,arrData)
End Function


'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetEnvironment()
		Dim conn,rs,strTemp
		On Error Resume NExt
		If objErr.Number=11 Then
			Exit Function
		End If

		set conn = CreateObject("ADODB.Connection")
		set rs = CreateObject("ADODB.Recordset")
	
		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &  gstrEnvironmentSetupDir & "\Setup.xls" & ";Excel 8.0;HDR=Yes;"  
		strQuery="Select *  From [Environment$]"
		rs.open strQuery,conn,1,1  
	
		If rs.RecordCount > 0 Then
				While not rs.EOF
					gstrEnv=Trim(rs.fields("Environment"))
					gstrExecuteMethod=rs.fields("ExecuteMethod")
					rs.moveNext
				Wend
		End If

		Set conn=Nothing
		Set rs=Nothing
		
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function Setup()
	Dim conn,rs,strTemp
	On Error Resume NExt
	If objErr.Number=11 Then
		Exit Function
	End If


	If gQCIntegration_Flag= True Then
		Set xlApp = CreateObject("Excel.Application")
		xlApp.Visible = False
		xlApp.DisplayAlerts = False
		Set xlWB = xlApp.Workbooks.Open(gstrEnvironmentSetupDir &"\Setup.xls")
		Set xlWS = xlWB.Worksheets("QCIntegration")

		For j= 0 To xlWS.UsedRange.Rows.Count-2
			ReDim Preserve gstrQCIntegrationArray(i)
			gstrQCIntegrationArray(i) = xlWS.Cells(i+2, 2).Value
			i=i+1
		Next
	
		xlWB.Close
		xlApp.Quit
		Set xlWB = Nothing
		Set xlApp = Nothing
	End If	

	Set dictApplicationURL=CreateObject("Scripting.Dictionary")
	set conn = CreateObject("ADODB.Connection")
	set rs = CreateObject("ADODB.Recordset")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" &  gstrEnvironmentSetupDir & "\Setup.xls" & ";Excel 8.0;HDR=Yes;"  
	strQuery="Select *  From [" & gstrEnv & "$]"
	rs.open strQuery,conn,1,1  

	If rs.RecordCount > 0 Then
		While not rs.EOF	   
			if Ucase(Trim(rs.fields("Brand")))="IBC" Then
				strIBCURL=Trim(rs.fields("URL"))
				Set gDictIBCData=CreateObject("Scripting.Dictionary")

				gDictIBCData.Add "IBCURL",strIBCURL
				gDictIBCData.Add "IBCUser",Trim(rs.fields("IBCUser"))
				gDictIBCData.Add "IBCUser1",Trim(rs.fields("IBCUser1"))
				gDictIBCData.Add "IBCDEV",Trim(rs.fields("IBCDEV"))
				gDictIBCData.Add  "IBCENV",Trim(rs.fields("IBCENV"))
				gDictIBCData.Add  "IBCTEST",Trim(rs.fields("System"))
			Else 
				 dictApplicationURL.Add Ucase(Trim(rs.fields("Brand"))),Trim(rs.fields("URL"))
				 'gstrEnvironment = rs.fields("System")
				 
			End If
			rs.moveNext
		Wend
	End If

	Set conn=Nothing
	Set rs=Nothing

End Function

'------------------------------------------------------------------------------------------------------------
'Function Name: updateGroupControlFile
'Description: To update execute flag as No
'------------------------------------------------------------------------------------------------------------
Function updateGroupControlFile(strStatus)

   Dim objExcel, objWrk, objSheet
   Dim nExecuteIndex, nTCIndex, nTCRowIndex

   Set objExcel = Createobject("Excel.Application")
   xlApp.Visible = False
   objExcel.DisplayAlerts = False
   Set objWrk = objExcel.Workbooks.Open(gstrControlFileName)
   Set objSheet = objExcel.WorkSheets(gstrGroupName)
   Set objSheetGroup = objExcel.WorkSheets("Groups")

	objSheet.Select
'	nExecuteIndex = objExcel.Application.WorksheetFunction.Match("Execute", objSheet.Rows("1:1"),0)
	nExecuteStatus = objExcel.Application.WorksheetFunction.Match("Status", objSheet.Rows("1:1"),0)
	nTCIndex = objExcel.Application.WorksheetFunction.Match("TestCaseID", objSheet.Rows("1:1"),0)
	nTCRowIndex = objExcel.Application.WorksheetFunction.Match(gstrCurScenario, objSheet.Columns(Chr(64 + nTCIndex) & ":" & Chr(64 + nTCIndex)), 0)
	nExecuteTime = objExcel.Application.WorksheetFunction.Match("Execution_Time", objSheet.Rows("1:1"),0)
'	objSheet.Cells(nTCRowIndex, nExecuteIndex).Value = strStatus
	objSheet.Cells(nTCRowIndex, nExecuteStatus).Value = strStatus
	objSheet.Cells(nTCRowIndex, nExecuteTime).Value = gResultLogtExeTime

	nTotalExecuteTime = objExcel.Application.WorksheetFunction.Match("Execution_Time", objSheetGroup.Rows("1:1"),0)
	nTotalTCIndex = objExcel.Application.WorksheetFunction.Match("Groups", objSheetGroup.Rows("1:1"),0)
	nTotalTCRowIndex = objExcel.Application.WorksheetFunction.Match(gstrGroupName, objSheetGroup.Columns(Chr(64 + nTotalTCIndex) & ":" & Chr(64 + nTotalTCIndex)), 0)
	objSheetGroup.Cells(nTotalTCRowIndex, nTotalExecuteTime).Value = gstrTotalExecutionTime
	
	objWrk.SaveAs gstrControlFileName
	objWrk.Close True
	objExcel.Quit
	Set objSheet = Nothing
	Set objWrk = Nothing
	Set objExcel = Nothing
   
End Function
'------------------------------------------------------------------------------------------------------------
