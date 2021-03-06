﻿'=====================================================================================================================
'Library File Name    :Report - SHARED DRIVE
'Author               :Sharad Mali
'Created date         :
'Description          :It lists the common Reports functions that can be used in the scripts.
'=====================================================================================================================
'01.	WriteHTMLHeader(ByRef clsEnvironmentVariables)
'02.	WriteHTMLErrorLog(ByRef clsEnvironmentVariables, ByVal rptDesc, ByVal rptShtNm, ByVal intRowNum, ByVal strInputValue)
'03.	WriteHTMLResultLog(ByVal rptDesc, ByVal intPassFail)
'04.	CreateReport(ByVal rptDesc, ByVal ExpectedResult, ByVal intPassFail)
'05.	Write_Summary_Header(ByVal Iteration_Count, ByRef objEnvironmentVariables)
'06.	WriteHTML_Verification()
'07.	WriteHTMLSummaryResultLog(ByVal rptDesc, ByVal intPassFail)
'08.	WriteHTML_Summary_Verification()
'09.	WriteTestResultInExcel (ByRef TestName , ByRef  StepsName ,ByRef stepDescription, ByRef ExpectedResult, ByRef  ActualResult, ByRef Status, ByRef UseExcelObject,ByRef ExcelFile)
'=====================================================================================================================

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTMLHeader
'Input Parameter    	:Environment variables
'Description        	:Function to create Run-Time Error Reporting HTML File on the 
'		  	 client machine & writes the header to it.
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------

'Option Explicit

Public Function WriteHTMLHeader(ByRef clsEnvironmentVariables)
	
	Dim fso, flReport, strFilePath
		
	strFilePath = clsEnvironmentVariables.RunTimeReportPath & "\" & gstrGroupName & "_" & gstrCurScenario & ".html"
		
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strFilePath)=FALSE Then
		strFileData = ""
		strFileData = strFileData + "<HTML>" + vbcrlf + vbtab + "<Title>Run-Time Error Reporting</Title>" + vbcrlf
		strFileData = strFileData + + vbtab + "<HEAD></HEAD>" + vbcrlf + vbtab +"<BODY>" + vbcrlf + vbtab + "<HR><Font Align=Center Name=CG Times 			Size=4 Style=Bold>" + vbcrlf + vbtab 
		Set flReport = fso.CreateTextFile(strFilePath,TRUE)
		flReport.WriteLine strFileData
	Else
		strFileData = ""
		strFileData = strFileData + vbcrlf + vbtab + "</TABLE>"
		strFileData = strFileData + vbcrlf + vbtab + "<BR><BR>"
		strFileData = strFileData + vbcrlf + vbtab + "<HR>"
		Set flReport = fso.OpenTextFile(strFilePath,2)
		flReport.WriteLine strFileData
	End If

	strFileData = ""
		
	strFileData = strFileData + "<TABLE frame=vsides Width=100%>" + vbcrlf + vbtab + vbtab + "<tr><th align=center>Execution Date/Time</th>" + vbcrlf + vbtab + vbtab + "<th align=center>Browser Name</th>" + vbcrlf + vbtab + vbtab + "<th align=center>Machine Name</th>" + vbcrlf + vbtab + vbtab + "<th align=center>Test Case Name</th>" + vbcrlf + vbtab + vbtab
	strFileData = strFileData + "<tr><td align=center>" + CStr(Now()) + "</td>" + vbcrlf + vbtab + vbtab + "<td align=center>" + clsEnvironmentVariables.BrowserName + "</td>" + vbcrlf + vbtab + vbtab + "<td align=center>" + Environment("LocalHostName") + "</td>" + vbcrlf + vbtab + vbtab + "<td align=center>" + gstrCurScenario + "</td></tr>" + vbcrlf + vbtab + "</TABLE>" + vbcrlf +vbtab + "<HR>"
	strFileData = strFileData + "</Font><br>" + vbcrlf + vbtab + "<font Name=Bodoni MT Style=Normal>" + vbcrlf + vbtab + "<TABLE Width=100% Border=1 bordercolor=#000000>"+ vbcrlf + vbtab
	strFileData = strFileData + "<Caption><h1><Font Name=CG Times Size=4 Style=Bold>The List of Failed Actions</Font></h1></Caption>" + vbcrlf + vbtab + vbtab + "<tr bgcolor=#D3D3D3>" + vbcrlf + vbtab+ vbtab +"<th>Sr.No.</th>" + vbcrlf + vbtab+ vbtab +"<th>Object Name</th>" + vbcrlf + vbtab+ vbtab +"<th>InputValue</th>" + vbcrlf + vbtab + vbtab +"<th>Failed Action</th>" + vbcrlf + vbtab + vbtab +"<th>Error Description</th>" + vbcrlf + vbtab + vbtab +"<th>Sheet Name</th>" + vbcrlf + vbtab +"</tr>" + vbcrlf + vbtab
		
	flReport.WriteLine strFileData
	flReport.Close
	Set flReport = Nothing
	Set fso = Nothing
		
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTMLErrorLog
'Input Parameter    	:Environment variables. report description
'Description        	:Function to open the already existing Run-Time Error Reporting
'		  	 HTML file whenever the error occurs & writes the details about
'		    	 the same error.
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteHTMLErrorLog(ByRef clsEnvironmentVariables, ByVal rptDesc, ByVal rptShtNm, ByVal intRowNum, ByVal strInputValue)
	Dim fso, flReport
	Dim rpt_ObjectName,rpt_InputValue,rpt_Action

	'rpt_ObjectName = DataTable("LabelName", rptShtNm)
	rpt_ObjectName = "TillDashboard"
	rpt_InputValue = strInputValue
'	rpt_Action = DataTable("Action", rptShtNm)
	rpt_Action = "TillDashboard"
	Test_Path = Environment("TestDir")
	Test_Script_Sheet=rptShtNm

	If rpt_ObjectName = "" Then
		rpt_ObjectName = Space(3)
	End If
	If rpt_InputValue = "" Then
		rpt_InputValue = "&nbsp"
	End If
	If rpt_Action = "" Then
		rpt_Action = Space(3)
	End If					
		
	strFilePath = clsEnvironmentVariables.RunTimeReportPath & "\" & gstrGroupName & "_" & gstrCurScenario & ".html"
		
	Set fso=Nothing
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set flReport = fso.OpenTextFile(strFilePath,8)

	strFileData = strFileData + "<tr>" + vbcrlf + vbtab+ vbtab +"<td>" + CStr(clsEnvironmentVariables.ErrNum) + "</td>" 		
	strFileData = strFileData + vbcrlf + vbtab + vbtab +"<td>" + rpt_ObjectName + "</td>" + vbcrlf + vbtab+ vbtab 		
	strFileData = strFileData + "<td>" + rpt_InputValue + "</td>" + vbcrlf + vbtab + vbtab +"<td>" + rpt_Action + "</td>" 		
	strFileData = strFileData + vbcrlf + vbtab + vbtab +"<td>" + rptDesc + "</td>" + vbcrlf + vbtab + vbtab +"<td>" 		
	strFileData = strFileData + "Row Number: " + CStr(intRowNum+1) + "<br>" 	
	strFileData = strFileData + "<a href="& chr(34) & clsEnvironmentVariables.ScriptPath & chr(34) &">"+ Test_Script_Sheet + "</a></td>" + vbcrlf + vbtab +"</tr>" + vbcrlf + vbtab		
	flReport.WriteLine strFileData		
	flReport.Close
	Set flReport = Nothing
	Set fso = Nothing
	
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTMLResultLog
'Input Parameter    	:rptDesc - The description from each function.
'                 	:intPassFail - If value of intPassFailis 0(zero) then action is failed for the keyword and
'                        will mark in red color  else if value is 1(One) then the action is Pass for
'                        the keyword and marked in GreenEnvironment variables. report description
'Description        	:This function will create the test Log File
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteHTMLResultLog_OLD(ByVal rptDesc, ByVal intPassFail)

	Dim fso, flReport, strFilePath, strFileData, strTime

	If gbReporting Then
	
		strFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".log"
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(strFilePath) = TRUE Then
				Set flReport = fso.OpenTextFile(strFilePath,8)
		Else
				Set flReport = fso.CreateTextFile(strFilePath,TRUE)
		End If
		If intPassFail = 0 Then	
				strFileData = strFileData & "<a href="& chr(34) & clsEnvironmentVariables.ScreenshotPath & chr(34) &"><b><font color=red>&nbsp;X FAIL </b></a>" & "&nbsp;</font>"
		ElseIf intPassFail = 1 Then
				strFileData = strFileData & "<b><font color=green><i>&radic;</i> PASS</font></b>" & "&nbsp;"		
		 ElseIf intPassFail = 4 Then
				strFileData = strFileData & "<a href="& chr(34) & clsEnvironmentVariables.ScreenshotPath & chr(34) &"><b><font color=green><i>&radic;</i> PASS</font></b></a>" & "&nbsp;"
		ElseIf intPassFail = 2 Then
				If InStr(rptDesc, "displayed successfully") > 0 Then
					strFileData = strFileData & "<br>"
				End If
				strFileData = strFileData & "<b><font color=blue><i> &nbsp;# </i> INFO </font></b>" & "&nbsp;"				 
		ElseIf intPassFail = 3 Then
				strFileData = strFileData & "<b><font color=orange><i> !! </i> WARNING</font></b>" & "&nbsp;"
		ElseIf intPassFail = 5 Then
				strFileData = strFileData & "&nbsp;"
		End If
	
		strFileData = strFileData & rptDesc & "<br>"
		flReport.Write strFileData
		flReport.Close
		Set flReport = Nothing
		Set fso = Nothing

	Else
		objErr.Clear
	End If

End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTMLResultLog
'Input Parameter    	:rptDesc - The description from each function.
'                 	:intPassFail - If value of intPassFailis 0(zero) then action is failed for the keyword and
'                        will mark in red color  else if value is 1(One) then the action is Pass for
'                        the keyword and marked in GreenEnvironment variables. report description
'Description        	:This function will create the test Log File
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteHTMLResultLog(ByVal rptDesc, ByVal intPassFail)

	Dim fso, flReport, strFilePath, strFileData, strTime
	If gbReporting Then
	
		strFilePath = gstrTestResultLogDir & gstrCurScenario & ".log"
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(strFilePath) = TRUE Then
				Set flReport = fso.OpenTextFile(strFilePath,8)
		Else
				Set flReport = fso.CreateTextFile(strFilePath,TRUE)
		End If
		If intPassFail = 0 Then	
'				gstrReportStatement = gstrReportStatement&"FAIL:-  " & rptDesc
'				gobjWord.Selection.TypeText(gstrReportStatement)
				Set objSelect = gobjWord.Selection
				strBookmarkName="p"&gstrScreenshotBookmarkNum
				gobjWord.Selection.Bookmarks.Add (strBookmarkName)
				set objImg=gobjWord.Selection.InlineShapes.AddPicture (gstrScreenShotPath,  True)
				objImg.ScaleHeight = 40
				objImg.ScaleWidth = 40
				gstrScreenshotBookmark="..\"&"Screenshots\"&gstrDocName&".doc"&"#"&strBookmarkName
				gstrScreenshotBookmarkNum=gstrScreenshotBookmarkNum+1
				objSelect.InsertBreak()
				gobjWord.ActiveDocument.Save		
				strFileData = strFileData & "<a href="& chr(34) & gstrScreenshotBookmark & chr(34) &"><b><font color=red>&nbsp;X FAIL </b></a>" & "&nbsp;</font>"
		ElseIf intPassFail = 1 Then
				strFileData = strFileData & "<b><font color=green><i>&radic;</i> PASS</font></b>" & "&nbsp;"		
		 ElseIf intPassFail = 4 Then
				Set objSelect = gobjWord.Selection
				strBookmarkName="p"&gstrScreenshotBookmarkNum
				gobjWord.Selection.Bookmarks.Add (strBookmarkName)
				set objImg=gobjWord.Selection.InlineShapes.AddPicture (gstrScreenShotPath,  True)
				objImg.ScaleHeight = 40
				objImg.ScaleWidth = 40
				gstrScreenshotBookmark="..\"&"Screenshots\"&gstrDocName&".doc"&"#"&strBookmarkName
				gstrScreenshotBookmarkNum=gstrScreenshotBookmarkNum+1
				objSelect.InsertBreak()
				gobjWord.ActiveDocument.Save
				Set  objImg=nothing
				Set objSelect=Nothing
				gstrReportStatement=" "
				strFileData = strFileData & "<a href="& chr(34) & gstrScreenshotBookmark & chr(34) &"><b><font color=green><i>&radic;</i> PASS</font></b></a>" & "&nbsp;"

		ElseIf intPassFail = 2 Then
				If InStr(rptDesc, "displayed successfully") > 0 Then
					strFileData = strFileData & "<br>"
				End If
'-----------------------------------------------------------------------------------------------------------------------------------------------------
'				Set objSelect = gobjWord.Selection
'				strBookmarkName="p"&gstrScreenshotBookmarkNum
'				gobjWord.Selection.Bookmarks.Add (strBookmarkName)
'				set objImg=gobjWord.Selection.InlineShapes.AddPicture (gstrScreenShotPath,  True)
'				objImg.ScaleHeight = 40
'				objImg.ScaleWidth = 40
'				gstrScreenshotBookmark="..\"&"Screenshots\"&gstrDocName&".doc"&"#"&strBookmarkName
'				gstrScreenshotBookmarkNum=gstrScreenshotBookmarkNum+1
'				objSelect.InsertBreak()
'				gobjWord.ActiveDocument.Save
'				Set  objImg=nothing
'				Set objSelect=Nothing
'				gstrReportStatement=" "
'				strFileData = strFileData & "<a href="& chr(34) & gstrScreenshotBookmark & chr(34) &"><b><font color=blue><i>&nbsp;#</i> INFO</font></b></a>" & "&nbsp;"
'-----------------------------------------------------------------------------------------------------------------------------------------------------
				strFileData = strFileData & "<b><font color=blue><i> &nbsp;# </i> INFO </font></b>" & "&nbsp;"			


					 
		ElseIf intPassFail = 3 Then
				strFileData = strFileData & "<b><font color=orange><i> !! </i> WARNING</font></b>" & "&nbsp;"
		ElseIf intPassFail = 5 Then
				strFileData = strFileData & "&nbsp;"
		End If
	
		strFileData = strFileData & rptDesc & "<br>"
		flReport.Write strFileData
		flReport.Close
		Set flReport = Nothing
		Set fso = Nothing

	Else
		objErr.Clear
	End If

End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:CreateReport
'Input Parameter    	:rptDesc - The description from each function.
'                 	:intPassFail - If value of intPassFailis 0(zero) then action is failed for the keyword and
'                        will mark in red color  else if value is 1(One) then the action is Pass for
'                        the keyword and marked in GreenEnvironment variables. report description
'Description        	:This function will create the .rpt File
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function CreateReport(ByVal rptDesc, ByVal intPassFail)

	Dim fso, flReport, strReportFilePath, strFileData, blnCreated,strExpectedResult
	gStrTestStep = gStrTestStep + 1
	strReportFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestresultLog\" & gstrCurScenario & ".rpt"

	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(strReportFilePath) = TRUE Then
      		Set flReport = fso.OpenTextFile(strReportFilePath,8)
      		blnCreated = True
	Else
      		Set flReport = fso.CreateTextFile(strReportFilePath,TRUE)
      		blnCreated = False
	End If

	If intPassFail = 0 Then
			
		' For character
		strFileData=""
		strFileData = strFileData & "FAIL "
		strExpectedResult= ExpectedResult
		If clsEnvironmentVariables.QCUpdation Then
			'Changes for uploading attachments in QC
			WriteTestResultInExcel gstrTestCaseNameFromQC, gStrTestStep ,gstrQCDesc,gstrExpectedResult, rptDesc,"FAILED",clsEnvironmentVariables.ScreenshotPath,clsEnvironmentVariables.UseExcelObject,clsEnvironmentVariables.TestResultExcelFile
		End If			
	Elseif intPassFail = 1 Then		
		' For Character
		strFileData=""
		strFileData = strFileData & "PASS "

		strExpectedResult= ExpectedResult		
		If clsEnvironmentVariables.QCUpdation Then
			WriteTestResultInExcel gstrTestCaseNameFromQC, gStrTestStep ,gstrQCDesc,gstrExpectedResult, rptDesc,"PASSED",clsEnvironmentVariables.ScreenshotPath,clsEnvironmentVariables.UseExcelObject,clsEnvironmentVariables.TestResultExcelFile
		End If		
	Elseif intPassFail = 2 Then
      		' For Character
		strFileData=""
		strFileData = strFileData & "INFO"
		strExpectedResult= ExpectedResult		
		If clsEnvironmentVariables.QCUpdation Then
			WriteTestResultInExcel gstrTestCaseNameFromQC, gStrTestStep ,gstrQCDesc,gstrExpectedResult, rptDesc,"PASSED","",clsEnvironmentVariables.UseExcelObject,clsEnvironmentVariables.TestResultExcelFile
		End If	
	Elseif intPassFail = 3 Then
       		'For Character
     		strFileData=""
      		strFileData = strFileData & "WARNING"
		strExpectedResult= ExpectedResult		
		If clsEnvironmentVariables.QCUpdation Then
			WriteTestResultInExcel gstrTestCaseNameFromQC, gStrTestStep ,gstrQCDesc,gstrExpectedResult, rptDesc,"PASSED","",clsEnvironmentVariables.UseExcelObject,clsEnvironmentVariables.TestResultExcelFile
		End If
		
    	End If
    	If blnCreated Then
      		strFileData = VbLf & strFileData & rptDesc
    	Else
      		strFileData = strFileData & rptDesc
    	End If
    	flReport.Write strFileData
    	flReport.Close
    	Set flReport = Nothing
    	Set fso = Nothing
	   clsEnvironmentVariables.ScreenshotPath=""
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:Write_Summary_Header
'Input Parameter    	:
'Description        	:This function will create the summary report heading
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function Write_Summary_Header(ByRef objEnvironmentVariables)
	
	Dim fso, MyFile
	Dim strLabel, strNewLabel, Flag, LenIndex
	
	strLabel = gstrGroupName
	strNewLabel = ""
	Flag = 0
	For LenIndex = 1 To Len(strLabel)
		
		If Mid(strLabel, LenIndex, 1) <> "" Then
			If Asc(Mid(strLabel, LenIndex, 1)) <= 91 And LenIndex > 1 Then
				Flag = LenIndex
				strNewLabel = strNewLabel & Left(strLabel, LenIndex - 1) & " "
				strLabel = Right(strLabel, Len(strLabel) - LenIndex + 1)
				LenIndex = 0
			End If
		End If
		
	Next
	
	strNewLabel = strNewLabel & strLabel
            
	If Flag = 0 Then
		strNewLabel = gstrGroupName
    End If
	
	strFilename = Environment.value("TestName")
	Set fso = CreateObject("Scripting.FileSystemObject")	
	Set MyFile = fso.CreateTextFile(gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName & "_" & rCurTime & "_Summary.html", True)
	MyFile.Close
	Set MyFile = fso.OpenTextFile(gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName & "_" & rCurTime & "_Summary.html",8)
	Myfile.Writeline("<html>")
	Myfile.Writeline("<head>")
	Myfile.Writeline("<meta http-equiv=" & "Content-Language" & "content=" & "en-us>")
	Myfile.Writeline("<meta http-equiv="& "Content-Type" & "content=" & "text/html; charset=windows-1252" & ">")
	Myfile.Writeline("<title>" & UCase(gstrProjectName) & " Automation Execution Summary Report</title>")
	Myfile.Writeline("</head>")
	Myfile.Writeline("<body>")
	Myfile.Writeline("<blockquote>")
	Myfile.Writeline("<p>")
	Myfile.Writeline("&nbsp;")
	Myfile.Writeline("</p>")
	Myfile.Writeline("</blockquote>")
	Myfile.Writeline("<p align=left>&nbsp;&nbsp;")
	Myfile.Writeline("<blockquote>")
	Myfile.Writeline("<blockquote>")
	Myfile.Writeline("<blockquote>")
	Myfile.Writeline("</p>")
	Myfile.Writeline("<table border=2 bordercolor=" & "#000000 id=table1 width=844 height=31 bordercolorlight=" & "#000000>")
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN =" & 3 & " bgcolor = #1E90FF>")
	Myfile.Writeline("<p><img src=" & chr(34) & gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\Logo\ClientLogo1.BMP"& chr(34) &" align =left><img src=" & chr(34) &  gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\Logo\ClientLogo.BMP"& chr(34) &" align =right></p><p align=center><font color=#000000 size=4 face= "& chr(34)&"Copperplate Gothic Bold"&chr(34) & ">&nbsp;" & UCase(gstrProjectName) & " Automation Execution Summary </font><font face= " & chr(34)&"Copperplate Gothic Bold"&chr(34) & "></font> </p>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = " & 3 & " bgcolor = #87CEFA>")
	Myfile.Writeline("<p align=justify><b><font color=#000000 size=2 face= Verdana>"& "&nbsp;"& "Scenario :&nbsp;&nbsp;" &  strNewLabel & "&nbsp;&nbsp;" & "</a>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = " & 3 & " bgcolor = #87CEFA>")
	Myfile.Writeline("<p align=justify><b><font color=#000000 size=2 face= Verdana>"& "&nbsp;"& "DATE :&nbsp;&nbsp;" &  now  & "&nbsp;&nbsp;" & "</a>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr>")
	Myfile.Writeline("<td COLSPAN = " & 3 & " bgcolor = #87CEFA>")
	Myfile.Writeline("<p align=justify><b><font color=#000000 size=2 face= Verdana>"& "&nbsp;"& "Executed By :&nbsp;&nbsp;" &   gstrTester  & "&nbsp;On&nbsp;" &  gstrMachine & "</a>")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr bgcolor=#F08080>")
	Myfile.Writeline("<td width=800")
	Myfile.Writeline("<p align=" & "left><b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">&nbsp;" & "Test Case Name</b></td>")
	Myfile.Writeline("<td width=100")	
	Myfile.Writeline("<p align=" & "left" & ">" & "<b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">&nbsp;" & "Result&nbsp;</b></td>")
	Myfile.Writeline("<td width=150")	
	Myfile.Writeline("<p align=" & "left" & ">" & "<b><font color = white face=" & "Arial Narrow " & "size=" & "2" & ">&nbsp;" & "Execution Time&nbsp;</b></td>")	
	Myfile.Writeline("</tr>")	
	Myfile.Writeline("</blockquote>")
	Myfile.Writeline("</body>")
	Myfile.Writeline("</html>")
	MyFile.Close

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTML_Summary_Verification
'Input Parameter    	:None
'Description        	:This function will create the .rpt File
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteHTML_Summary_Verification()
    	Dim fso, flReport, flLog, strFilePath, strLogFilePath, strRepFilePath, strTime, strFileData, strVerification, timeInMinute, timeInSecond, tExeTime
    	Dim cn, rs, strGExecutionTime
    	Dim rTime

    	rTime = date
    	rTime= replace(rTime,"/","_")
    	gSummaryreport="Summary_Scenario_TestLog" & rTime&".html"
    	strFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName & "_" & rCurTime & "_Summary.html"
    	strLogFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName & "_" & rCurTime & "_Summary.log"	

    	Set fso = CreateObject("Scripting.FileSystemObject")
    	If fso.FileExists(strLogFilePath) = False Then
      		Exit Function
    	End If
    	Set flLog = fso.OpenTextFile(strLogFilePath,1)    
      	Set flReport = fso.OpenTextFile(strFilePath,8)
    	
    	strFileData = flLog.ReadAll
	flLog.Close    	
    	flReport.Write strFileData
    	flReport.Close
	
  	Set flLog = fso.GetFile(strLogFilePath)
  	flLog.Delete(True)
  	Set flLog = Nothing
	strRowCount = Datatable.GetSheet(gstrGroupName).GetRowCount
	
	timeInMinute =gGroupExecutionTime \ 60
	timeInHours= timeInMinute \ 60
	timeInMinute = timeInMinute - (60 * timeInHours)
    timeInSecond=  (gGroupExecutionTime) mod 60
	strGExecutionTime = "Total Execution Time is " & timeInHours & " hours " & timeInMinute & " minute " & timeInSecond & " seconds"
	gstrTotalExecutionTime=timeInHours& " : " & timeInMinute & ": " & timeInSecond
	Set Myfile = fso.OpenTextFile(gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName &  "_" & rCurTime &"_Summary.html",8)
	Myfile.Writeline("<tr bgcolor = #FDEEF4 ><td COLSPAN =3 bgcolor = white><b><font color=#000080 face=Verdana size=2>&nbsp;" & strGExecutionTime & "</b></font></td></tr>")
	Myfile.Writeline("<tr bgcolor = #FDEEF4>")
	Myfile.Writeline("<td width = 300")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& "Total Count&nbsp;&nbsp;")
	Myfile.Writeline("</td>")
	Myfile.Writeline("<td COLSPAN =2 width = 300>")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& strRowCount & "&nbsp;")
	Myfile.Writeline("</td>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr bgcolor = #FDEEF4>")
	Myfile.Writeline("<td width = 300>")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& "Total Executed&nbsp;&nbsp;")
	Myfile.Writeline("</td>")
	Myfile.Writeline("<td COLSPAN =2 width = 300>")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& gTotalCount  & "&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("</tr>")
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr bgcolor = #FDEEF4>")
	Myfile.Writeline("<td width = 300>")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& "Total Passed&nbsp;&nbsp;")
	Myfile.Writeline("</td>")
	Myfile.Writeline("<td COLSPAN =2 width = 300>")
	Myfile.Writeline("<p align=justify><b><font color=green size=2 face= Verdana>"& "&nbsp;"& gTotalPassed  & "&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr bgcolor = #FDEEF4>")
	Myfile.Writeline("<td width=300>")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& "Total Failed&nbsp;&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("<td COLSPAN =2 width=300>")
	Myfile.Writeline("<p align=justify><b><font color=red size=2 face= Verdana>"& "&nbsp;"& gTotalFailed  & "&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("</tr>")
	Myfile.Writeline("<tr bgcolor = #FDEEF4>")
	Myfile.Writeline("<td width=300>")
	Myfile.Writeline("<p align=justify><b><font color=#000080 size=2 face= Verdana>"& "&nbsp;"& "Total No Run&nbsp;&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("<td COLSPAN =2 width=300>")
	Myfile.Writeline("<p align=justify><b><font color=gray size=2 face= Verdana>"& "&nbsp;"& (strRowCount-gTotalCount)  & "&nbsp;")
	Myfile.Writeline("</td>")	
	Myfile.Writeline("</tr>")
	Myfile.Writeline("</table>")
	Myfile.Close				

	'SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe", gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName &  "_" & rCurTime &"_Summary.html"
	gstrCurSummaryPath = gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName &  "_" & rCurTime &"_Summary.html"
  	Call UpdateSummaryExecutionResultLog(gstrTotalExecutionTime)
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTML_Verification
'Input Parameter    	:None
'Description        	:Function to create Verification Result HTML File on the
'             		 central machine & writes the header to it if it does not exists.
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteHTML_Verification()
    	Dim fso, flReport, flLog, strFilePath, strLogFilePath, strRepFilePath, strTime, strFileData, strVerification, timeInMinute, timeInSecond, tExeTime
    	Dim cn, rs,strTemp
    	Dim rTime
    	rTime = date & "_" & time
    	rTime= replace(rTime,"/","")
    	rTime= replace(rTime,":","")
		rTime= replace(rTime," ","")

		strLogFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".log"
    	strRepFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario & ".rpt"

		Set fso = CreateObject("Scripting.FileSystemObject")
    	If fso.FileExists(strLogFilePath) = False Then
      		Exit Function
    	End If
    	Set flLog = fso.OpenTextFile(strLogFilePath,1)
		strTemp = flLog.ReadAll
		
		If InStr(strTemp,"X FAIL") > 0 Then
			strFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario &"_"& rTime &"_Log_Fail.html"
			gdetailReport=strFilePath	
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Word Screenshot
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            gobjDoc.SaveAs gstrScreenshotDocPath&"_FAIL.doc" ,0
			testResult="FAIL"
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		Else
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			gobjDoc.SaveAs gstrScreenshotDocPath &"_PASS.doc" ,0
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			strFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\TestResultLog\" & gstrCurScenario &"_" & rTime &"_Log_Pass.html"
			gdetailReport=strFilePath	
		End If
		
		timeInMinute=ExecutionTime \ 60
		timeInSecond=(ExecutionTime) mod 60
		tExeTime=timeInMinute &" Minute "& timeInSecond &" Second"
		gResultLogtExeTime=timeInMinute &":"& timeInSecond 
	
    	If fso.FileExists(strFilePath)=FALSE Then		

			strFileData = strFileData + "<html><head><meta http-equiv=Content-Languagecontent=en-us><meta http-equiv=Content-Typecontent=text/html; charset=windows-1252><title>" + gstrProjectName + " Automation Execution Test Log</title></head>"
			strFileData = strFileData + "<body><blockquote><p>&nbsp;</p></blockquote><p align=left>&nbsp;&nbsp;<blockquote><blockquote><blockquote></p>"
			strFileData = strFileData + "<table border=2 bordercolor=#000000 id=table1 width=844 height=31 bordercolorlight=#000000>"
			strFileData = strFileData + "<tr><td COLSPAN =3 bgcolor = #2F755D><p><img src="& chr(34) & gstrSummaryReportDir +"\Logo\ClientLogo1.BMP"& chr(34) &" align =left><img src=" & chr(34) & gstrBaseDir &"\Reports\" & gstrProjectUser & "\Summary\Logo\ClientLogo.BMP"& chr(34) &" align =right></p><p align=center><font color=white size=4 face= "& chr(34)&"Copperplate Gothic Bold"& chr(34)&">&nbsp;" & gstrProjectName & " Automation Test Log </font><font face= "& chr(34)&"Copperplate Gothic Bold"& chr(34)&"></font> </p></td></tr>"
			strFileData = strFileData + "<tr><td width=300 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;Test Case ID</a></td>"
			strFileData = strFileData + "<td width=800 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;"+ gstrCurScenario +"</a></td></tr>"
			strFileData = strFileData + "<tr><td width=300 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;Description</a></td>"
			strFileData = strFileData + "<td width=800 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;" + gstrcurScenarioDesc + "</a></td></tr>"
			strFileData = strFileData + "<tr><td width=300 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;Execution Time</a></td>"
			strFileData = strFileData + "<td width=800 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;" + tExeTime + "</a></td></tr>"

			strFileData = strFileData + "<tr><td width=300 bgcolor = #ccffff><p align=justify><b><font color=#000000 size=2 face= Verdana>&nbsp;Status</a></td>"
					If InStr(strTemp,"X FAIL") > 0 Then
						strFileData = strFileData + "<td width=800 bgcolor = #ccffff><p align=justify><b><font color=red size=2 face= Verdana>&nbsp;FAIL</a></td></tr>"
					else
						strFileData = strFileData + "<td width=800 bgcolor = #ccffff><p align=justify><b><font color=green size=2 face= Verdana>&nbsp;PASS</a></td></tr>"
					End If 
			
			strFileData = strFileData + "<tr><td COLSPAN =3 bgcolor = #2F755D><p align=center><b><font color=white size=2 face= Verdana>&nbsp;Results</a></td></tr>"
			
            Set flReport = fso.CreateTextFile(strFilePath,TRUE)
      		flReport.Write strFileData			
    	Else	
			Set flReport = fso.OpenTextFile(strFilePath,2)
    	End If

		strFileData = ""
		strFileData = strFileData & "<tr><td COLSPAN =3 bgcolor = #ccffff><font face=Arial size=2><br>"
		strFileData = strFileData & strTemp  
		strFileData = strFileData & "&nbsp;</td></tr>"
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Word Screenshot
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        gstrFileData=strFileData	
		If testResult="FAIL" Then
					strFileData=replace(gstrFileData,".doc","_Fail.doc",1,-1,0)
			else 
					strFileData=replace(gstrFileData,".doc","_Pass.doc",1,-1,0)
		End If
		gobjDoc.close
		fso.deletefile(gstrScreenshotDocPath&".doc")	

		If fso.FileExists(gstrScreenshotDocPath&".doc")=TRUE Then		
					fso.DeleteFolder(gstrScreenshotsDir & gstrCurScenario )
		End If

		Set gobjWord=nothing
		Set gobjDoc=nothing

        Set Wshell=CreateObject("WScript.Shell")
		Set oExe=Wshell.Exec("taskkill /F /IM winword.exe")
		Set oExe= Nothing
		Set Wshell= Nothing
		
'		systemUtil.CloseProcessByName("WINWORD.EXE")
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		flReport.Write strFileData
		flLog.Close
		Set flLog = fso.GetFile(strLogFilePath)
		flLog.Delete(True)
		Set flLog = Nothing
		Set flLog = fso.OpenTextFile(strRepFilePath,1)
		strVerification = flLog.ReadAll
		Set fso=Nothing
		flLog.Close
	
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteHTMLSummaryResultLog
'Input Parameter    	:rptDesc -
'        		 intPassFail -if value of intPassFailis 0(zero) then action is failed for the keyword and will mark in red color
'                     	 else if value is 1(One) then the action is Pass for the keyword and marked in Green
'Description        	:This function will create the test Log File
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteHTMLSummaryResultLog(ByVal intPassFail)

    	Dim fso, flReport, strFilePath, strFileData, strTCExecutionTime

    	strFilePath = gstrBaseDir & "Reports\" & gstrProjectUser & "\Summary\" & gstrGroupName & "_" & rCurTime  & "_Summary.log"
    
		Set fso = CreateObject("Scripting.FileSystemObject")
    	If fso.FileExists(strFilePath) = TRUE Then
      		Set flReport = fso.OpenTextFile(strFilePath,8)
    	Else
      		Set flReport = fso.CreateTextFile(strFilePath,TRUE)
    	End If	
	timeInMinute=ExecutionTime \ 60
	timeInHours = timeInMinute \ 60
    timeInSecond=(ExecutionTime) mod 60	
	strTCExecutionTime = timeInHours & ":" & timeInMinute & ":" & timeInSecond
  	If intPassFail=3 then				
		strFileData = strFileData & "<tr bgcolor = #FDEEF4 ><td width=300><p align=left><font face=Verdana size=2>&nbsp;" & gstrCurScenario & "</a></td><td width=300nowrap><p align=justify><b><a href=" & chr(34) & gdetailReport & chr(34) & " target=_blank><font face=Verdana size=2 color=green>&nbsp;WARNING</font></b></td><td width=100><font face=Verdana size=2>&nbsp;" & strTCExecutionTime & "</font></b></td></tr>"
   	ElseIf intPassFail = 0 Then        	
		strFileData = strFileData & "<tr bgcolor = #FDEEF4 ><td width=300><p align=left><font face=Verdana size=2>&nbsp;" & gstrCurScenario & "</a></td><td width=300nowrap><p align=justify><b><a href=" & chr(34) & gdetailReport & chr(34) & " target=_blank><font face=Verdana size=2 color=red>&nbsp;FAIL</font></b></td><td width=100><font face=Verdana size=2>&nbsp;" & strTCExecutionTime & "</font></b></td></tr>"
    	ElseIf intPassFail = 1 Then
		strFileData = strFileData & "<tr bgcolor = #FDEEF4 ><td width=300><p align=left><font face=Verdana size=2>&nbsp;" & gstrCurScenario & "</a></td><td width=300nowrap><p align=justify><b><a href=" & chr(34) & gdetailReport & chr(34) & " target=_blank><font face=Verdana size=2 color=green>&nbsp;PASS</font></b></td><td width=100><font face=Verdana size=2>&nbsp;" & strTCExecutionTime & "</font></b></td></tr>"        	
    	End If    	
    	flReport.Write strFileData
    	flReport.Close
    	Set flReport = Nothing
    	Set fso = Nothing
	Call UpdateExecutionResultLog(gstrCurScenario,intPassFail, strTCExecutionTime,gstrDataReusabilityStatus )
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:WriteTestResultInExcel
'Input Parameter    	:None
'Description        	:This function will write test result in excel sheet
'Calls              	:errorHandler in case of any error
'Return Value       	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WriteTestResultInExcel(ByRef TestName , ByRef  StepsName ,ByRef stepDescription, ByRef ExpectedResult, ByRef  ActualResult, ByRef Status, BYRef screenshotpath,ByRef UseExcelObject,ByRef ExcelFile)

	'Create an Excel Object		

	' Set the active sheet
	Set objSheet = UseExcelObject.ActiveSheet
  
	' Set the sheet name'		
	objSheet.Name = "AutoTest"

	'Set the columen header of excel
	objSheet.Cells(1, 1).Value = "TestCase_Name"
	objSheet.Cells(1, 2).Value = "Step_Name"
	objSheet.Cells(1, 3).Value = "Description"
	objSheet.Cells(1, 4).Value = "Expected_Result"
	objSheet.Cells(1, 5).Value = "Actual_Result"
	objSheet.Cells(1, 6).Value = "Status"
	objSheet.Cells(1, 7).Value = "Screenshot_Path"

	Set rngUsed = objSheet.UsedRange

	nRows=rngUsed.Rows.count
	nRow = nRows+1
		
	' Set the excel values
	objSheet.Cells(nRow, 1).Value = TestName
	objSheet.Cells(nRow, 2).Value = StepsName
	objSheet.Cells(nRow, 3).Value =stepDescription
	objSheet.Cells(nRow, 4).Value = ExpectedResult
	objSheet.Cells(nRow, 5).Value = ActualResult
	objSheet.Cells(nRow, 6).Value = Status	
		objSheet.Cells(nRow, 7).Value = screenshotpath	
	UseExcelObject.ActiveWorkbook.SaveAs(ExcelFile)

End Function
'---------------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
'Function Name       : PageInformation
'Input Parameter     : strVal - Text to print in report
'Description         : This function prints text in the report. Here it is used to display the Page Name in report file.
'------------------------------------------------------------------------------------------------------------
Public Function pageInformation(strVal) 
    Dim bResult,strDesc
    bResult=True  
    gstrDesc ="The Page with name '" & strVal & "' is displayed."
    WriteHTMLResultLog gstrDesc, 2
'    CreateReport  gstrDesc, 1  
    PageInformation = bResult
End Function  
'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:UpdateExecutionResultLog
'Input Parameter    	:strCurScenario - TC Name
'        		 				  intStatus - TC Execution status
'								  strExecutionTime - TC execution time.
'Description        	   :This function is used to updated TCs exeuction status in "Execution.mdb"
'Return Value       	  :None
'---------------------------------------------------------------------------------------------------------------------
Function UpdateExecutionResultLog(Byval strCurScenario, Byval intStatus, Byval strExecutionTime,Byval gstrDataReusabilityStatus)

	Dim strExecutionFileName, strExecutionStatus, strGroupName,strIteration

	If gstrDataReusabilityStatus ="EXECUTEDATATESTCASE" and gstrDataTestCaseIteration=""Then
		gstrDataTestCaseIteration=2
		Exit Function
	End If

		If Cint(intStatus) = 0 Then
			intStatus = "Fail"
			strExecutionStatus = "Y"
		ElseIf Cint(intStatus) = 1 then
			intStatus = "Pass"
			strExecutionStatus = "Y"  'Previous val = "E"
		ElseIf Cint(intStatus) = 3 Then
			intStatus = "Warning"
			strExecutionStatus = "Y"  'Previous val = "E"
		End If

		Set con = CreateObject("ADODB.Connection")
		Datafile=gstrControlFilesDir & "\GroupControlFiles" & gstrProjectUser & ".xls"
		con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="  & Datafile & ";Extended Properties=""Excel 8.0""" 
		strSQL= "Update ["& gstrGroupName &"$] Set [Status]= '"& intStatus &"', [Execute]= '"& strExecutionStatus &"', [Execution_Time]= '"& strExecutionTime &"' where TestCaseID= '"& strCurScenario &"'"
		con.Execute strSQL
		con.Close
		Set con=Nothing
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:UpdateExecutionResultLog
'Input Parameter    	:strCurScenario - TC Name
'        		 				  intStatus - TC Execution status
'								  strExecutionTime - TC execution time.
'Description        	   :This function is used to updated TCs exeuction status in "Execution.mdb"
'Return Value       	  :None
'---------------------------------------------------------------------------------------------------------------------
Function UpdateSummaryExecutionResultLog(Byval strExecutionTime)


		Set cn = CreateObject("ADODB.Connection")
		Set rs = CreateObject("ADODB.Recordset")

		Datafile=gstrControlFilesDir & "\GroupControlFiles" & gstrProjectUser & ".xls"
		cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="  & Datafile & ";Extended Properties=""Excel 8.0""" 

		strSQL= "Update [Groups$] Set [ExecutionTime]= '"& strExecutionTime &"' where Groups= '"& gstrGroupName &"'"
		cn.Execute strSQL
		cn.Close
		Set cn=Nothing
End Function
'---------------------------------------------------------------------------------------------------------------------
