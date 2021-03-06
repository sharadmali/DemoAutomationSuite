﻿'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'Library File Name  	: InitScript
'Author             	: Sharad Mali
'Created date       	: 
'Modified date        	:  
'Description        	: It lists the common functions which are used before initializing a test run
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'*********************************************************************************************
'Function Name		    :ReadGroupData
'Input Parameter    	:None, but The ADDRESS of an array which is required to store the Group Names is passed
'Description        	:Reads the Groupfile Name from database table Groups and
											'stores all the group file name in arrGroup array
'Calls              	:
'Return Value	:
'*********************************************************************************************
Public Function ReadGroupData(byref arrGroup)
	Dim nCounter
	Dim con 
	Dim rs 
	Dim str 
	nCounter = 0
	Set con = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")
	str = "SELECT * from [Groups$] where Execute='Y'"    ' Query to select groups of each suite 	
	con.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & gstrControlFilesDir & "GroupControlFiles" & gstrProjectUser & ".xls"
	rs.Open str, con
	rs.MoveFirst	
	While(not rs.EOF)  		
		If UCASE(rs("Execute").Value) = "Y" then
					redim preserve arrGroup(nCounter)
					arrGroup(nCounter) = rs("Groups").Value   ' Group names being fed into the array
					nCounter=nCounter+1	
		End if   		  		
		rs.movenext
	wend	
	rs.close
	con.close
	set rs=nothing
							
End Function

'-------------------------------------------------------------------------------------------------
'Function Name		: getIniFileData
'Input Parameter    	: None
'Description        	: Reads the data in the ini file and deletes the ini file
'Calls              	:
'Return Value	        :
'-------------------------------------------------------------------------------------------------
Public Function getIniFileData()

	Dim objFile, ObjTextFile, ObjTextFileOpen, nLinesCount
	Dim strTempFileName
	strTempFileName = strBaseDir & "Config\Test.ini"
	
	'Create a File System Object
	Set objFile = CreateObject("Scripting.FileSystemObject")

	'Get the Ini file passed by the web page for reading	
	Set ObjTextFile			= objFile.GetFile(strBaseDir & "\Config\Test.ini")
	Set ObjTextFileOpen = ObjTextFile.OpenAsTextStream

	gstrEnvironment			= ObjTextFileOpen.ReadLine()
	gstrEntity					= ObjTextFileOpen.ReadLine()
	gstrImpactUser			= ObjTextFileOpen.ReadLine()
	gstrImpactPassword	= ObjTextFileOpen.ReadLine()
	gstrTranDate				= ObjTextFileOpen.ReadLine()
	gnTestDay						= ObjTextFileOpen.ReadLine()
	gstrRunName					= ObjTextFileOpen.ReadLine()

	'Close the Ini file
	ObjTextFileOpen.Close

	'Delete the Ini file
	objFile.DeleteFile(strTempFileName)

	' Free the memory
	Set objFile = Nothing

End Function

'------------------------------------------------------------------------------------------------------------
'Function Name 		:executedatatestcase
'Input Parameter	:None
'Description		:This function executes data test case
'Calls			:None
'Return	Value		:None
'------------------------------------------------------------------------------------------------------------
Public Function executedatatestcase(intDataID,strDatasheet,strBrand,UpLoadExcelReport)
   Dim nTempColCount, conn, rs, fso, strQuery, nLoopCount, strTemp, arrTemp
   Dim nTempIndex, arrTempAction, arrTempDataSet, objFile, arrCustomer

	set conn = CreateObject("ADODB.Connection")
	set rs = CreateObject("ADODB.Recordset")
	
	
	gstrRig = gDictEnvironment("RIG")
	gstrtestenvironment = gDictEnvironment("TAROT")
	
	If strBrand="HAL" Then
		strBrand="Halifax"
	End If

	If gstrDataReusabilityStatus="NONREUSABLE"  Then
		strQuery = "SELECT * FROM NewUsers Where BRAND = '" & strBrand &"' and  TESTINGSERVER= '" &gstrRig&"'and STATUS= '"&struserstatus&"'"
		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb;User Id=;Password=;"		
		rs.open strQuery,conn,1,1
		intCount = rs.RecordCount 
		
		Set conn=Nothing
		Set rs=Nothing
		Set fso=Nothing
		
		If intCount >0 Then
			Exit Function
		End If
	End If
		
	If intCount=0 OR gstrDataReusabilityStatus="EXECUTEDATATESTCASE" Then	
		set conn = CreateObject("ADODB.Connection")
		set rs = CreateObject("ADODB.Recordset")
		Set fso = CreateObject("Scripting.FIleSystemObject")
'		str= "SELECT * from ["&strDatasheet&"$] where ID='TC_SIT_E2E_002'"
		'str2 = "SELECT * from [Groups$] where Execute='Y'" 
		str = "SELECT * from ["&strDatasheet&"$] where TestCaseID='"&intDataID&"'"

		conn.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & gstrControlFilesDir & "\GroupControlFiles" & gstrProjectUser & ".xls"
		rs.Open str, conn
		
		gstrCurScenario = rs( "TestCaseID")
		gstrcurScenarioDesc = rs( "Description")
		gstrCurBatch = rs("Batch_Test_File")
		gstrTestCaseNameFromQC = rs( "TestCaseNameFromQC")
		gstrHostName =rs( "Host_Name")
		ExcelFilePath= gstrBaseDir &"Reports\" & gstrProjectUser & "\TestResultLog\"
		ExcelFileWithPath = ExcelFilePath & "TestResults.xls"
		
		If (fso.FileExists(ExcelFileWithPath)) Then
			fso.deletefile(ExcelFileWithPath)
		End If
		
		Set conn=Nothing
		Set rs=Nothing
		Set fso=Nothing
		
		WriteHTMLHeader clsEnvironmentVariables  	
		strStartTime=Left( MonthName( Month( Date() ) ), 3 ) & " " & Day( Date() ) & ", " & Year( Date() ) & " " & Time()   
		  If  blnQCUpdation Then
				clsEnvironmentVariables.TestResultExcelFile=ExcelFileWithPath
				Set clsEnvironmentVariables.UseExcelObject =  clsQCIntegrationModule.CreateExcelFile()
		End If
		
		bresult = DriverScript		
		
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
	
	'gstrStrUsername=""
	gstrtemppswd=""
	strtempusname=""

	
			
End Function
'------------------------------------------------------------------------------------------------------------


