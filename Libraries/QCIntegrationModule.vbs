﻿'==================================================================================================================
'Library File Name    :QCIntegrationModule - SHARED DRIVE
'Author               :Sharad Mali
'Created date         :
'Description          :It lists the QC Integration functions that can be used in the scripts.
'==================================================================================================================

Class QCIntegrationModule	


	Public Excel


	'---------------------------------------------------------------------------------------------------------------------
	'Function Name		:DownloadAttachementFromQC
	'Input Parameter	:ByVal FolderName,ByVal TestNameFromQC,ByVal DownloadPath,ByVal TestScriptName
	'Description		:This function is used to download attachment attached to test case of qc
	'Calls          	:None
	'Return Value 		:None
	'---------------------------------------------------------------------------------------------------------------------
	Function DownloadAttachementFromQC(ByVal FolderName, ByVal TestNameFromQC,ByVal DownloadPath,ByVal TestScriptName)    

		Set testFactory= QCConnection.TestFactory			

		If Len(FolderName) <> 0 Then
			SubjectPath = "Subject\" & FolderName
		Else
			SubjectPath = "Subject" 
		End If
			
		Set TestFilter = testFactory.Filter 
		TestFilter.Filter("TS_NAME") = Chr(34) & Trim(TestNameFromQC)  & Chr(34)
		Set TestList = testFactory.NewList(TestFilter.Text)
		Set GetTest = TestList.Item(1)
		Set attachFact = GetTest.Attachments 
		Set attachList = attachFact.NewList("") 
						
		For each tAttach In attachList
			Set attachemntstorage = tAttach.AttachmentStorage
			
			attachemntstorage.ClientPath = DownloadPath 						
			QCFileName=tAttach.name(0)
			ActualFileName=tAttach.name(1)		
			DownloadFileName =Mid(ActualFileName,1,Len(ActualFileName)-4)
										
			If StrComp(TestScriptName,DownloadFileName,1)=0 Then									
				attachemntstorage.Load tAttach.name,True
				Set renfile = CreateObject("Scripting.FileSystemObject")												
				If renFile.FolderExists(attachemntstorage.ClientPath) Then
					If  renFile.FileExists(attachemntstorage.ClientPath & "\" & ActualFileName) Then
						Set delfile = renFile.GetFile(attachemntstorage.ClientPath & "\" & ActualFileName)
						delfile.delete
					End If									
					renFile.MoveFile attachemntstorage.ClientPath & "\" & QCFileName,attachemntstorage.ClientPath & "\" & ActualFileName 
				End If
			End If
			
		Next

	End Function
	'---------------------------------------------------------------------------------------------------------------------

	'---------------------------------------------------------------------------------------------------------------------
	'Function Name		:UploadTestResultsInQCFromExcel
	'Input Parameter	:ByVal spa_Path, ByVal spa_tst ,ByVal  ExcelFile
	'Description		:This function uploads the test result in QC
	'Calls          	:None
	'Return Value 		:None
	'---------------------------------------------------------------------------------------------------------------------
	Function UploadTestResultsInQCFromExcel(ByVal spa_Path, ByVal spa_tst ,ByVal  ExcelFile)
	   'msgbox "Inside UploadTestResultsInQCFromExcel"
		Dim StartLoopCnt,MaxLoopCnt
		Dim blnNextTestCase,blnFailed, TestCaseFoundFlag 
		
		TestCaseFoundFlag = False
		GetFromExcel TestNameArr, StepsNameArr, DescriptionArr, ExpectedResultArr, ActualResultArr, StatusArr, AttachmentPathArr, ExcelFile

		MaxLoopCnt= UBound(TestNameArr)	
		blnProceed = True			
		
		StartLoopCnt = 0
		
		Set spa_TrMgr = QCConnection.TestSetTreeManager		
		If isEmpty(spa_TrMgr) or  spa_TrMgr is nothing Then
				Exit function
		End If
		Set spa_Folder = spa_TrMgr.NodeBypath(spa_Path)	
		If isEmpty(spa_Folder) or  spa_Folder is nothing Then
				Exit function
		End If
		Set spa_TestSetF = spa_Folder.TestSetFactory		
		Set spa_tsetList = spa_TestSetF.NewList("")
		'msgbox spa_tsetList
		For Each spa_Mytestset In spa_tsetList
		   
			If spa_Mytestset.Name = spa_tst Then
				 'msgbox  spa_Mytestset.Name &"   "&spa_tst&"Inside first if"
				Set spa_TestSetF1  = spa_MYtestset.TSTestFactory
				Set spa_TestSetF2  = QCConnection.TestFactory
				blnNextTestCase = True
				StartLoopCnt = 0 
				'msgbox StartLoopCnt&"     "&MaxLoopCnt
				Do while StartLoopCnt <= MaxLoopCnt
					spa_Tid= TestNameArr(StartLoopCnt)	   
					'msgbox   spa_Tid   		
					If blnNextTestCase Then
							Set spa_Testsetlist = spa_TestSetF2.Newlist("SELECT  *  FROM  TEST  WHERE  TS_NAME ='" & spa_Tid & "'")
							str_Test_Case_Flag=True
							 For each spa_Tstt IN spa_TestsetList
									str_Test_Case_Flag=False
									Set aFilter = spa_TestSetF1.Filter
									aFilter.Filter("TC_TEST_ID") = spa_Tstt.ID
									Set lst = spa_TestSetF1.NewList(aFilter.Text)
									If lst.Count = 0 Then
										Set spa_MYtestset1 = spa_TestSetF1.AddItem(spa_Tstt.ID)  												   
										spa_TestsetID = spa_Tstt.ID
									Else
										Set spa_MYtestset1 = lst.Item(1)
									End If
									If  spa_MYtestset1.islocked Then
										spa_MYtestset1.UnLockObject
									End If

									spa_TestsetID = spa_Tstt.ID
									'spa_Mytestset1.Status= "No Run" ' spa_Status
									spa_Mytestset1.Post

									Set spa_RunF = spa_Mytestset1.RunFactory
									Set spa_Myrun= spa_RunF.AddItem("Run_" & DatePart("m",Date) & "-" & DatePart("d", Date) & "_" & DatePart("h",Time) & "-" & DatePart("n", Time)  & "-" & DatePart("s", Time))
																	
									spa_Myrun.Status = "No Run"'spa_Status
									spa_Myrun.Post	  										   
                                 
									If str_Test_Case_Flag=False Then
									   Exit For
									End If

								Next
					End If
					
						StepName = StepsNameArr(StartLoopCnt)					
						StepDescription = DescriptionArr(StartLoopCnt)
						ExpectedResult =ExpectedResultArr(StartLoopCnt)
						ActualResult =ActualResultArr(StartLoopCnt)
						Status=StatusArr(StartLoopCnt)
						AttachmentPath = AttachmentPathArr(StartLoopCnt)
											
						Set spa_run = spa_Myrun.StepFactory
						Set spa_Mystep= spa_run.AddItem(StepName)
						spa_Mystep.Name = StepName
						spa_Mystep.Field("ST_DESCRIPTION")=StepDescription
						spa_Mystep.Field("ST_EXPECTED")=ExpectedResult
						spa_Mystep.Field("ST_ACTUAL")=ActualResult
						spa_Mystep.Status = Status
						spa_Mystep.Post
	
						Set spa_MystepAtt = spa_Mystep.Attachments
						Set spa_MystepAttach = spa_MystepAtt.AddItem(Null) 
						'msgbox AttachmentPath
						If  Trim(AttachmentPath) <> "" And UCase(Status) = "PASSED" Then
							'msgbox "Add Attachment"
							spa_MystepAttach.FileName = AttachmentPath
							spa_MystepAttach.Type = 1
							spa_MystepAttach.Description="Passed ScreenShot"
							spa_MystepAttach.Post
						
						End If 
				   
						If  Trim(AttachmentPath) <> "" And Status = "FAILED" Then
							'msgbox "Add Attachment"
							spa_MystepAttach.FileName = AttachmentPath
							spa_MystepAttach.Type = 1
							spa_MystepAttach.Description="Failed ScreenShot"
							spa_MystepAttach.Post

						End If 
					'End If
                                        
					StartLoopCnt = StartLoopCnt+1
					If  (StrComp(ucase(Status),"FAILED",1) = 0)  Then
						blnFailed =True
					End If								
					If StartLoopCnt <= MaxLoopCnt  Then
						If  (StrComp(spa_Tid, TestNameArr(StartLoopCnt),1) <>0)  Then
							blnNextTestCase = True
						Else
							blnNextTestCase = False
						End If
					Else
						blnNextTestCase = True
					End If										
					If blnNextTestCase Then
						If blnFailed Then
							spa_Myrun.Status = "Failed"
						Else
							spa_Myrun.Status = "Passed"
						End If
						spa_Myrun.Post
						blnFailed = False
					End If									
				loop
			End If
		Next
	Set spa_RunF=Nothing
		'Attach Detailed File to Test Run
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set fldr = fso.GetFolder(gstrTestResultLogDir)
		Set flsc = fldr.Files
		arrLatestFile=split(gdetailReport,"\")
		
		For Each fl In flsc
			If Instr(fl.Name,arrLatestFile(Ubound(arrLatestFile))) Then
				Set spa_run_Attach = spa_Myrun.Attachments
				Set spa_run_Attach_Null = spa_run_Attach.AddItem(fl.Name)
				spa_run_Attach_Null.Description = "Detailed File"
				spa_run_Attach_Null.post
				Set o_ExtStr = spa_run_Attach_Null.AttachmentStorage
				o_ExtStr.ClientPath = gstrTestResultLogDir & "\" 
    				o_ExtStr.Save fl.Name, true
				spa_run_Attach_Null.post
				Exit For
			End If
		Next

		Set fso = CreateObject("Scripting.FileSystemObject")
		Set fldr = fso.GetFolder(gstrScreenshotsDir)
		Set flsc = fldr.Files
		arrLatestFile=split(gstrScreenshotDocPath,"\")
		
		For Each fl In flsc
			If Instr(fl.Name,arrLatestFile(Ubound(arrLatestFile))) Then
				Set spa_run_Attach = spa_Myrun.Attachments
				Set spa_run_Attach_Null = spa_run_Attach.AddItem(fl.Name)
				spa_run_Attach_Null.Description = "Detailed File"
				spa_run_Attach_Null.post
				Set o_ExtStr = spa_run_Attach_Null.AttachmentStorage
				o_ExtStr.ClientPath = gstrScreenshotsDir & "\" 
    				o_ExtStr.Save fl.Name, true
				spa_run_Attach_Null.post
				Exit For
			End If
		Next
'		If FlgStatus = False Then
'			Set fldr = fso.GetFolder(gstrScreenshotsDir)
'			Set flsc = fldr.Files
'			For Each fl In flsc
'				If Instr(fl.Name,gstrCurScenario) Then
'					Set spa_run_Attach = spa_Myrun.Attachments
'					Set spa_run_Attach_Null = spa_run_Attach.AddItem(fl.Name)
'					spa_run_Attach_Null.Description = "Failed Snapshot"
'					spa_run_Attach_Null.post
'					Set o_ExtStr = spa_run_Attach_Null.AttachmentStorage
'    					o_ExtStr.ClientPath = gstrScreenshotsDir & "\"
'    					o_ExtStr.Save fl.Name, true
'					spa_run_Attach_Null.post
'				End If
'			Next	
'		End If 
	End Function
	'---------------------------------------------------------------------------------------------------------------------

	'---------------------------------------------------------------------------------------------------------------------
	'Function Name		:GetFromExcel
	'Input Parameter	:ByRef TestName , ByRef  StepsName ,ByRef stepDescription, ByRef ExpectedResult, ByRef  ActualResult,
	'			 ByRef Status, ByVal ExcelFile
	'Description		:This function gets the results stored in temporary Excel file
	'Calls          	:None
	'Return Value 		:None
	'---------------------------------------------------------------------------------------------------------------------
	Function GetFromExcel (ByRef TestName , ByRef  StepsName ,ByRef stepDescription, ByRef ExpectedResult, ByRef  ActualResult, ByRef Status, ByRef AttachmentPath, ByVal ExcelFile)
        	Dim Sheet,usedRowsCount,rowObj,columnObj,curRow,curCol,strTestID,strTestName 
        	Dim strStepname,strActualResult,strDescription,strExpectedResult,strAttachmentPath
		    Set ExcelObject = CreateObject("Excel.Application")
        	ExcelObject.DisplayAlerts = False
        	ExcelObject.Workbooks.Open(ExcelFile)
	   'msgbox TestName
		Set Sheet = ExcelObject.Sheets("AutoTest")
		'Set Sheet = ExcelObject.Sheets(TestName)
        
        	'Total rows are used in the current worksheet
        	usedRowsCount =  Sheet.UsedRange.Rows.Count 
        	columnObj = 1

		For rowObj = 2 To (usedRowsCount)
			curRow = rowObj 
			curCol = columnObj 
			'get the value that is in the cell 		
			strTestName =  strTestName &  Sheet.Cells(curRow, curCol).Value & "|"
			strStepname = strStepname &  Sheet.Cells(curRow, curCol + 1).Value & "|"
			strDescription = strDescription &  Sheet.Cells(curRow, curCol + 2).Value & "|"
			strExpectedResult = strExpectedResult &  Sheet.Cells(curRow, curCol + 3).Value & "|"
			strActualResult = strActualResult & Sheet.Cells(curRow, curCol+4).Value & "|"
			strStatus = strStatus & Sheet.Cells(curRow, curCol+5).Value & "|"
			strAttachmentPath = strAttachmentPath & Sheet.Cells(curRow, curCol+6).Value & "|"
        	Next
        	strTestName = Mid(strTestName,1,Len(strTestName) -1)  
        	strStepname = Mid(strStepname,1,Len(strStepname) -1) 
			strDescription =Mid(strDescription,1,Len(strDescription) -1) 
			strExpectedResult =Mid(strExpectedResult,1,Len(strExpectedResult) -1) 
        	strActualResult = Mid(strActualResult,1,Len(strActualResult) -1) 
			strStatus = Mid(strStatus,1,Len(strStatus) -1) 
			strAttachmentPath = Mid(strAttachmentPath,1,Len(strAttachmentPath) -1)
		
		'Split the string into array to fill  the value into array
		TestName = Split(strTestName,"|")
		StepsName = Split(strStepname,"|")
		StepDescription = Split(strDescription,"|")
		ExpectedResult = Split(strExpectedResult,"|")
		ActualResult = Split(strActualResult,"|")
		Status = Split(strStatus,"|")
		AttachmentPath = Split(strAttachmentPath,"|")

        	If Not Sheet Is Nothing Then
			set  Sheet = Nothing
		End If		
		If Not ExcelObject Is Nothing Then
			ExcelObject.Quit()
			set ExcelObject = Nothing
		End If
	End Function
	'---------------------------------------------------------------------------------------------------------------------

	'---------------------------------------------------------------------------------------------------------------------
	'Function Name		:CloseExcelFile
	'Input Parameter	:Excel object
	'Description		:This function will close the excel instance
	'Calls          	:None
	'Return Value 		:None
	'---------------------------------------------------------------------------------------------------------------------
	Function CloseExcelFile(ByVal ObjExcel)		
		ObjExcel.Quit
		set ObjExcel = Nothing
	End Function
	'---------------------------------------------------------------------------------------------------------------------

	'---------------------------------------------------------------------------------------------------------------------
	'Function Name		:CreateExcelFile
	'Input Parameter	:None
	'Description		:This function will create the new excel file
	'Calls          	:None
	'Return Value 		:None
	'---------------------------------------------------------------------------------------------------------------------
	Function CreateExcelFile()
	
		 'MsgBox" Create an Excel Object"
		Set ObjExcel = CreateObject ("Excel.Application")
	
		'Add workbook
		ObjExcel.WorkBooks.Add()
	
		' Set the alerts to false for excel
		ObjExcel.DisplayAlerts = 0	
	
		Set CreateExcelFile = ObjExcel
	End Function

	'---------------------------------------------------------------------------------------------------------------------
	'Function Name		:UploadResultLogInQC
	'Input Parameter	:ByVal strDriveName, ByVal strProjectName
	'Description		:This function will upload execution logs in QC
	'Calls          	:None
	'Return Value 		:None
	'---------------------------------------------------------------------------------------------------------------------
	Function UploadResultLogInQC(ByVal strDriveName, ByVal strProjectName)    			
			msgbox "UploadResultLogInQC"
        	QCProject =QCSuitePath & SuiteName
		
		TestResultLogPathInQC = QCProject & "Reports\TestResultLog"
		TestResultLogPath = gstrTestResultLogDir 
		ScreenshotPathInQC =  QCProject & "Reports\ScreenShots"
		ScreenshotPath = gstrScreenshotsDir
   
		'Set QC = QCUtil.QCConnection
		'Set treeMgr = QC.TreeManager
 	
		Set attachFolder = treeMgr.NodeByPath(TestResultLogPathInQC)
		'Call the FolderAttachment routine to upload all the attachment	
		Set Attachments = attachFolder.Attachment
		'Get the list of attachments and upload new attachement one at a time. 
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set fldr = fso.GetFolder(TestResultLogPath)
		Set flsc = fldr.Files
		For Each fl In flsc
			If UCase(Right(fl.Name,4)) = "HTML" Then
				Rep_File_Path= TestResultLogPath & "\" & fl.Name
				arrTemp = Split(fl.Name,".",-1,1)
				strFix = Replace(Replace(Now,"/",""),":","")
				strFix = Replace(strFix," ","")				
				Rep_NewFile_Path = TestResultLogPath & "\" & arrTemp(0) & strFix & "." & arrTemp(1)
				fl.Copy Rep_NewFile_Path							
				Set Attachment = Attachments.AddItem(Null)
				Attachment.FileName = Rep_NewFile_Path
				Attachment.Description = "Test Result Log "
				Attachment.Type = 1
				Attachment.Post 
				Wait(2)
				Set fld = fso.GetFile(Rep_NewFile_Path)
				Set fld1 = fso.GetFile(Rep_File_Path)
				fld.Delete
				Set fld = Nothing
				Set fld1 = Nothing
			End If			
		Next 				
		TestSummaryResultLogPathInQC=QCProject & "\Reports\Summary"
		TestSummaryResultLogPath = gstrSummaryReportDir 
		Set attachFolder = treeMgr.NodeByPath(TestSummaryResultLogPathInQC)
		'Call the FolderAttachment routine to upload all the attachment	
		Set Attachments = attachFolder.Attachments
		'Get the list of attachments and upload new attachement one at a time. 
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set fldr = fso.GetFolder(TestSummaryResultLogPath )
		Set flsc = fldr.Files
		For Each fl In flsc
			If UCase(Right(fl.Name,4)) = "HTML" Then
				Rep_File_Path= TestSummaryResultLogPath & "\" & fl.Name
				arrTemp = Split(fl.Name,".",-1,1)
				strFix = Replace(Replace(Now,"/",""),":","")
				strFix = Replace(strFix," ","")				
				Rep_NewFile_Path = TestSummaryResultLogPath & "\" & arrTemp(0) & strFix & "." & arrTemp(1)
				fl.Copy Rep_NewFile_Path						
	
				Set Attachment = Attachments.AddItem(Null)
				Attachment.FileName = Rep_NewFile_Path
				Attachment.Description = "Test Summary Result Log "
				Attachment.Type = 1
				Attachment.Post 
				Wait(2)
				Set fld = fso.GetFile(Rep_NewFile_Path)
				Set fld1 = fso.GetFile(Rep_File_Path)
				fld.Delete
				Set fld = Nothing
				Set fld1 = Nothing
			End If			
		Next 	
		Set attachFolder = treeMgr.NodeByPath(ScreenshotPathInQC)
		' Call the FolderAttachment routine to upload all the attachment	
		Set Attachments = attachFolder.Attachments
		' Get the list of attachments and upload new attachement one at a time. 		
		Set fldr = fso.GetFolder(ScreenshotPath)
		Set flsc = fldr.Files
		For Each fl In flsc
			If UCase(Right(fl.Name,3)) = "DOC" Then
				Rep_File_Path= ScreenshotPath & "\" & fl.Name				
				Set Attachment = Attachments.AddItem(Null)
				Attachment.FileName = Rep_File_Path
				Attachment.Description = "Screen shot"
				Attachment.Type = 1
				Attachment.Post
			End If 	
			If UCase(Right(fl.Name,3)) = "PNG" Then
				Rep_File_Path= ScreenshotPath & "\" & fl.Name				
				Set Attachment = Attachments.AddItem(Null)
				Attachment.FileName = Rep_File_Path
				Attachment.Description = "Screen shot"
				Attachment.Type = 1
				Attachment.Post
			End If 				
		Next 
		Set fl = Nothing
		Set flsc = Nothing
		Set fldr = Nothing
		Set fso = Nothing
		Set QC = Nothing
		
	End Function
	Public Function ConnectToQC(byVal  strURL,byVal strUserName,byVal strPassword,byVal strDomain,byVal strProject)
		Set QCConnection=CreateObject("TDApiOle80.TDConnection")
		If QCConnection.Connected Then
			If QCConnection.LoggedIn  Then
				QCConnection.Logout
			End If
			QCConnection.Disconnect
		End If
		QCConnection.InitConnectionEx  strURL 
		QCConnection.Login  strUserName, strPassword
		QCConnection.Connect strDomain,strProject
		Set treeMgr  = QCConnection.TreeManager			
	End Function

	Public Function DisConnectQC()
		If QCConnection.Connected Then
			If QCConnection.LoggedIn  Then
				QCConnection.Logout
			End If
			QCConnection.Disconnect
		End If
		Set QCConnection=Nothing
	End Function
	'---------------------------------------------------------------------------------------------------------------------

End Class
