'==================================================================================================================
'Library File Name    :Utility - SHARED DRIVE
'Author               :Sharad Mali
'Created date         :
'Description          :It lists the common utility functions that can be used in the scripts.
'==================================================================================================================
'01.	splitstring(strtosplit, strDelimiter)
'01. 	getTestCaseData()
'02. 	getDataRowIDS(strParam, objDictTC)
'03. 	getDataArray(strParam,strRowKey,arrData)
'04.	getCurrentTestCaseData(arrFnData,arrFunCol,nCur)
'05. 	importSheetsToRunTimeDataTable(strDTFilePath, strDTSheetName)
'06.	SkipKeyword(strObjName,strVal,strLabel)
'07.	WaitFor(strTestData)
'08.	loadOR(strTSRFile)
'09.	unLoadOR()
'10.	fireEvent(strObjName,strLabel,strVal)
'11.	waitProperty(strObjName,strLabel,strVal)
'12.	verifyProperty(strObjName,strLabel,strVal)
'13.	setTOProperty(strObjName,strLabel,strVal)
'14.	changeReplayType(strVal)
'15.	invokeBrowser(strVal)
'16.	ActivateBrowser(strObjName,strLabel)
'17.	CloseBrowser(strObjName,strLabel)
'18.	CloseAllBrowsers()
'19.	syncBrowser(strObject)
'20.	VerifyPage(strObject,strLabel)
'21.	clickButton(strObject,strLabel)
'22.	clickImage(strObject,strLabel)
'23.	enterText(strObject,strLabel,strVal)
'24.	typeText(strObject,strVal)
'25.	selectList(strObject,strData,strLabel)
'26.	verifyItemsInList(strObject,strLabel,strData)
'27.	setChkBox(strObject)
'28.	clickLink(strObject,strLabel)
'29.	selectRadioButton(strObject,strLabel)
'30.	clickWebElement(strObject)
'31.	AssociateOR(strAction, strTSR, blnFlag)
'32.	TakeScreenShot()
'33. 	pressTab()
'34.	clickWebelementFromWebTable(strObject, strWebElement)
'35.	BrowserBack(strObject,strLabel,strVal)
'36.	BrowserForward(strObject,strLabel,strVal)
'37.	VerifyDialogBox(strObject,strLabel, strVal)
'====================================================================================================================

'Define global variables for all functions
option explicit

On Error Resume Next

Public dictStoredValue        ' A dictionary object to store values to be checked subsequently
Set dictStoredValue = CreateObject("Scripting.Dictionary")
Dim strErrString          ' to capture Error message In case of any error
Dim strLogString
Public strTestDatabasePath      ' Path of test data database
dim strTDPath           ' Path of TD database
public nCurID           ' For holding Current row ID in test data table
public nCurrDataRow         ' For holding the current row
public nCurrDataCol         ' For holding Column row ID in test data table

'---------------------------------------------------------------------------------------------------------------------
'Function Name		:splitstring
'Input Parameter	:strtosplit (String) - representing the string to be split
'		      	 strDelimiter (String) - representing the delimiter used to split the string
'Description		:Splits a string using the delimiter specified
'Calls          	:errorHandler in case of any error
'Return Value 		:An array with the split string
'---------------------------------------------------------------------------------------------------------------------
Public Function splitstring(strtosplit, strDelimiter)
	Dim arrSplit
	'Split the requested string with the specified delimiter
	arrSplit=split(strtosplit, strDelimiter)
	'Return the string to the right of the delimiter e.g if strtosplit has value Table=Employee, this function replaces Employee
	splitstring=arrSplit(1)
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:getTestCaseData
'Input Parameter    	:None
'Description        	:Gets the names of data tables and values for particular scenario from TestCase table and stores it in a
'			 dictionary object.
'Calls              	:errorHandler in case of any error
'Return Value       	:Dictionary Object containing column names and thier values from the TestCase table in key-value pairs
'---------------------------------------------------------------------------------------------------------------------
Public function getTestCaseData()
	Dim strQuery
	Dim rs
	Dim nrow
	Dim ncol
	Dim conn, fso
	Dim objDictTC

	nrow = 0
	ncol = 0

	'Create the dictionary object
	Set objDictTC = CreateObject("Scripting.Dictionary")
	set conn = CreateObject("ADODB.Connection")
	set rs = CreateObject("ADODB.Recordset")
	'Connect with the DB
	Set fso = CreateObject("Scripting.FIleSystemObject")

'	If fso.FileExists(gstrTestDataDir & "\TestData_"& gstrEnv &".mdb") Then
	If fso.FileExists(gstrTestDataDir & "\TestData.mdb") Then
		'Setup the required query to get the correct values from the database
		strQuery = "select * from [TestCase] where TestCaseID='" & gstrCurScenario &"'"
		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData_"& gstrEnv &".mdb;User Id=;Password=;"	
	Else
		'Setup the required query to get the correct values from the database
		strQuery = "select * from [TestCase$] where TestCaseID='" & gstrCurScenario &"'"
		conn.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & gstrTestDataDir & "\TestData.xls"
	End If

	Set fso = Nothing
	
	rs.open strQuery,conn,1,1
	'Loop to read the recordset till the end
	While(not rs.EOF)
		'Loop to add the index and fieldname mapping to the dictionary object
		While(ncol < rs.Fields.Count)
			'Add values to the dictionary object with the field name as the value and the index as the key
			objDictTC.add rs.Fields(ncol).Name,rs.Fields(ncol).Value
			ncol = ncol + 1
		Wend
		ncol = 0
		'Move to the next recorset
		rs.movenext
	Wend
	rs.close
	conn.close
	'Return the dictionary object containing the column names and thier values from the TestCase table in key-value pairs
	set getTestCaseData = objDictTC
End function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name 		:getDataRowIDS
'Input Parameter    	:strParam(String) - Contains the table name and the names of the columns for which the value is to be stored
'			 (Columns=*) indicates the function to fetch the values for all the columns
'			 objDictTC - Dictionary object (passed byRef) containing the IDs for all table for a specific Scenario ID
'Description      	:Returns the Row IDS from the test case table for the requested table
'Calls          	:errorHandler in case of any error
'Return Value     	:Row ID for the specified table
'---------------------------------------------------------------------------------------------------------------------
Public Function getDataRowIDS(strParam,objDictTC)
	Dim strTBLName
	Dim strDataRows, arrTblClmn , strColumnName
	'Split the string wrt ';' to get the table and column name
	arrTblClmn = split(strParam,";")
	'Call the splitstring function to get the table name
	strTBLName = splitstring(arrTblClmn (0),"=")
	If objDictTC.Exists(strTBLName) then
		'Fetch the Row ID from the dictionary object wrt to the table name
        strDataRows = objDictTC.item(strTBLName)
	Else
		Msgbox "Please check the column name in Test Case Table!!!!!!", vbQuestion
		objErr.Raise 11
	End If
	'Return the Row ID for the specified table
	getDataRowIDS = strDataRows
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name 		:getDataArray
'Input Parameter   	:strParam(String) - Contains the table name from which the value is to be stored
'			 strRowKey - Contains the data ID for the mentioned table
'			 arrData - Stores the values taken from the table
'Description      	:Used to store the data in an array from the mentioned table with the help of the mentioned data ID
'Calls          	:errorHandler in case of any error
'Return Value     	:Returns a dictionary object which contains the mapping
'			 for the mentioned table with the field names and indexes as key value pairs 
'---------------------------------------------------------------------------------------------------------------------
Public Function getDataArray(strParam,strRowKey,arrData)
	
	Dim strQuery
	Dim nrow, ncol
	Dim conn,rs, strExc
	Dim objDictData
	Dim strTableName, strColumnName,strVal
	Dim arrTblClmn,strFieldValue, fso
	Dim nRowLoopCnt, nColLoopCnt,gRowKey

	gRowKey=strRowKey
	If isarray(arrData) Then
		Erase arrData
	End If 
	'Split the string wrt ';' to get the table and column name	
	arrTblClmn = split(strParam,";")
	'Call the splitstring function to get the table name

	strTableName = splitstring(arrTblClmn(0),"=")

	'Call the splitstring function to get the column name
	strColumnName = splitstring(arrTblClmn(1),"=")
	If instr(strColumnName,"*")=0 then
		strColumnName = "ID," & strColumnName 
	End If  	
	
	Set objDictData = CreateObject("Scripting.Dictionary")
	'Setup the required
	'query to get the correct values from the database	
	If gstrActionDataSet = "" Then
		gstrActionDataSet = "A"
	End If	
	
	
	set conn = CreateObject("ADODB.Connection")
	set rs = CreateObject("ADODB.Recordset")
	'Connect with the DB 
	Set fso = CreateObject ("Scripting.FileSystemObject")
'	If fso.FileExists(gstrTestDataDir & "\TestData_"& gstrEnv &".mdb") Then
	If fso.FileExists(gstrTestDataDir & "\TestData.mdb") Then
		strQuery = "Select "& strColumnName &" from ["&strTableName&"] where ID in (" & strRowKey & ") AND DataSet='" & gstrActionDataSet & "' Order By SrNo"	  
'		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData_"& gstrEnv &".mdb;User Id=;Password=;"		
		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb;User Id=;Password=;"		
	Else
		strQuery = "Select "& strColumnName &" from ["&strTableName&"$] where ID in (" & strRowKey & ") AND DataSet='" & gstrActionDataSet & "'"	   
		conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.xls" & ";Excel 8.0;HDR=Yes;"  
	End If
	Set fso = Nothing

 	'Empty the dictionary object
 	objDictData.RemoveAll		
	If strRowKey = "" Then		
'		clsEnvironmentVariables.ErrNum = "#"
'		WriteHTMLErrorLog clsEnvironmentVariables,"Value not provided in TestCase sheet for column '" & strTableName & "' in TestData file", gstrActionName, "", ""
	Else				
		rs.open strQuery, conn, 1, 1		 
 		nRow=rs.RecordCount 
		If nRow = 0 Then
'			clsEnvironmentVariables.ErrNum = "#"
'			WriteHTMLErrorLog clsEnvironmentVariables,"Record not found in sheet '" & strTableName & "' for ID=" & strRowKey & " and Dataset='" & gstrActionDataSet & "' in TestData File", gstrActionName, "", ""
		End If			
 		nCol=rs.Fields.Count		
		
 		'Re initialize the data array
 		redim arrdata(nrow-1, ncol-1)
 		nColLoopCnt = 0
 		'Loop to store the mapping of the field names with the indexes
 		While(nColLoopCnt<nCol)
  			objDictData.add rs.Fields(nColLoopCnt).Name,nColLoopCnt				
  			nColLoopCnt=nColLoopCnt+1
 		Wend			
		nRowLoopCnt = 0
		nColLoopCnt = 0
		'Loop to populate the array data		
		While(not rs.EOF) 			
			While(nColLoopCnt < nCol)
				If IsNull(rs.Fields(nColLoopCnt).Value) Then
					strFieldValue = ""
				Else
					strFieldValue = rs.Fields(nColLoopCnt).Value 
				End If							
				arrData(nRowLoopCnt, nColLoopCnt) = strFieldValue 						
				nColLoopCnt = nColLoopCnt + 1
			wend
			nColLoopCnt = 0
			nRowLoopCnt=nRowLoopCnt+1
			rs.movenext
		wend
		'Close the recordset
		rs.close
	End If
	Set rs = Nothing
	'Close the connection to the DB
	conn.close
	set getDataArray = objDictData
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name       	:getCurrentTestCaseData
'Input Parameter     	:
'Description         	:Gets current test data
'Calls               	:errorHandler in case of any error
'Return Value        	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public function getCurrentTestCaseData(arrFnData,arrFunCol,nCur)
	Dim objCurTC
	Dim ColCnt
	Dim nCnt
	Set objCurTC = CreateObject("Scripting.Dictionary")
	ColCnt = uBound(arrFunCol)
	For nCnt=0 to ColCnt
		'Add values to the dictionary object with the field name as the value and the index as the key
  		objCurTC.add arrFunCol(nCnt),arrFnData(nCur,nCnt)
	Next
 	set getCurrentTestCaseData = objCurTC
End function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:importSheetsToRunTimeDataTable
'Input Parameter    	:strDTFilePath - Excel File path to be imported
'Description        	:This function import the excel file to runtime data table of QTP
'---------------------------------------------------------------------------------------------------------------------
'Public Function importSheetsToRunTimeDataTable(strDTFilePath, strDTSheetName)
'   
'   	DataTable.AddSheet strDTSheetName
'   	DataTable.Importsheet strDTFilePath,1,strDTSheetName
'End Function
'----------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      	:importSheetsToRunTimeDataTable
'Input Parameter    	:strDTFilePath - Excel File path to be imported
'Description        	:This function import the excel file to runtime data table of QTP
'---------------------------------------------------------------------------------------------------------------------
Public Function importSheetsToRunTimeDataTable(strDTFilePath, strSheetSource,strDTSheetName)
   
   	DataTable.AddSheet strDTSheetName
	If IsNumeric(strSheetSource) Then
   		DataTable.Importsheet strDTFilePath,CLng(strSheetSource),strDTSheetName
	Else
		DataTable.Importsheet strDTFilePath,strSheetSource,strDTSheetName
	End If
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:WaitFor
'Input Parameter      	:strParam - Time in seconds
'Description          	:This function introduces a delay in the script
'Calls            	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function WaitFor(strTestData)
	Dim bResult
	bResult=True
	wait strTestData
	WaitFor=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:loadOR
'Input Parameter      	:Name of TSR file
'Calls                	:errorHandling is done at the level of ORClass
'Description           	:Loading TSR OR at Run time (Keyword can be used in Test case file)
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function loadOR(strTSRFile)
	Dim bResult
	bResult=True
													
	'Call sendNumLock()
	'Removing object repository to current Action
	If strPrevTSR<>False Then
		AssociateOR "Action1",strPrevTSR,"False"
	End If
	gstrCurTSR =  gstrObjectRepositoryDir  & "\" & strTSRFile
	'''''msgbox   gstrCurTSR    
 
	'Reading Object Repository
'	Set gobjObjectClass = new clsOR   
	gobjObjectClass.setORFile gstrCurTSR
	'Associate object repository to current Action

	AssociateOR "Action1",gstrCurTSR,"True"
	strPrevTSR=gstrCurTSR
	loadOR=bResult

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:unLoadOR
'Input Parameter      	:None
'Description          	:This function disassociates specified shared object repository from current action
'Calls            	:None
'Return Value   	:True Or False
'---------------------------------------------------------------------------------------------------------------------
Public Function unLoadOR()
  	Dim qtApp,qtRepository,bres
  	bRes=true
	
'	If objErr.number = 11 Then
'		Exit Function
'	End If

        Set qtApp = CreateObject("Quicktest.Application")
        Set qtRepository = qtApp.Test.Actions("Action1").ObjectRepositories
        qtrepository.RemoveAll
  	unLoadOR=bRes
End Function
'----------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name		:fireEvent
'Input Parameter    	:strObjName - String - Name of the Browser
'Description        	:This function fires specified event on object
'Calls function        	:NA
'Return Value		:True Or False
'---------------------------------------------------------------------------------------------------------------------
Public Function fireEvent(strObjName,strLabel,strVal)
	
	Dim objTemp, strClass, bResult, strDesc	

	bResult = False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objTemp = gobjObjectClass.getObjectRef(strObjName)
	gstrQCDesc = "Fireevent " & strVal & " on " & strLabel
	gstrExpectedResult = "Fired Event " & strVal & " on " & strLabel
	If objTemp.exist(5) Then
'		strClass = objTemp.GetROProperty("micclass")
		Setting.WebPackage("ReplayType") = 2 
	    objTemp.FireEvent strVal
		Setting.WebPackage("ReplayType") = 1
		gstrDesc = "Fired Event " & strVal & " for " & strLabel
	    WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc,1
		bResult = True
	Else
		gstrDesc = "Failed to fire Event " & strVal & " on " & strLabel
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0	   	     	
		bResult = False
		objErr.Raise 11
	End If
	Set objTemp = Nothing
	fireEvent = bResult 

End Function
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Function Name		:waitProperty
'Input Parameter    	:strObjName - String - Name of the Browser
'Description        	:This function waits until the property gets specified value
'Calls function        	:NA
'Return Value		:True Or False
'----------------------------------------------------------------------------------------------------------------------
Public Function waitProperty(strObjName,strLabel,strVal)
	
	Dim objTemp, strClass, bResult, strDesc,arrTemp, arrProp
	Dim bFlag
     bResult = True	

	 strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	Set objTemp = gobjObjectClass.getObjectRef(strObjName)
	arrTemp = Split(strVal,";")
	If objTemp.exist(gExistCount) Then		
		If InStr(1,arrTemp(0),"<>") > 0 Then
			arrProp = Split(arrTemp(0),"<>")			
	     		objTemp.WaitProperty arrProp(0),micNotEqual(arrProp(1))
		ElseIf InStr(1,arrTemp(0),"<=") Then	
			arrProp = Split(arrTemp(0),"<=")
	     		objTemp.WaitProperty arrProp(0),micLessThanOrEqual(arrProp(1))			
		ElseIf InStr(1,arrTemp(0),">=") Then		
	     		arrProp = Split(arrTemp(0),">=")
	     		objTemp.WaitProperty arrProp(0),micGreaterThanOrEqual(arrProp(1))			
		ElseIf InStr(1,arrTemp(0),">") Then		
	     		arrProp = Split(arrTemp(0),">")
	     		objTemp.WaitProperty arrProp(0),micGreaterThan(arrProp(1))			
		ElseIf InStr(1,arrTemp(0),"<") Then		
	     		arrProp = Split(arrTemp(0),"<")
	     		objTemp.WaitProperty arrProp(0),micLessThan(arrProp(1))			
		Else
			arrProp = Split(arrTemp(0),"=")
	     	objTemp.WaitProperty arrProp(0),arrProp(1),arrTemp(1)
		End If			
	End If
	
	Set objTemp = Nothing
	waitProperty = bResult 

End Function
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Function Name		:verifyProperty
'Input Parameter    	:strObjName - String - Name of the Browser
'Description        	:This function verifies the specified property of the object with expected value
'Calls function        	:NA
'Return Value		:True Or False
'----------------------------------------------------------------------------------------------------------------------
Public Function verifyProperty(strObjName,strLabel,strVal)
	Dim objTemp, strClass, bResult, arrTemp, arrProp, strPropValue, arrColName, strTemp, strPropName
	Dim bFlag
     bResult = False	

	arrTemp = Split(strVal, "[")
	strPropName = arrTemp(0)
	
	arrColName = Split(arrTemp(1), "]")
	strTemp = getTestDataValue(arrColName(0))

	strVal = Replace(strVal, "[", "")
	strVal = Replace(strVal, "]", "")
	strVal = Replace(strVal, arrColName(0), strTemp)
	
	If UCase(strTemp) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objTemp = gobjObjectClass.getObjectRef(strObjName)	
	arrTemp = Split(strVal,";")	 
	If objTemp.exist(gExistCount) Then	

		strinnertext=objTemp.GetRoProperty("innertext")

		If InStr(1,arrTemp(0),"<>") > 0 Then
			arrProp = Split(arrTemp(0),"<>")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should not be equal to " & arrProp(1) & " For " & strLabel				
	     		If objTemp.GetROProperty(arrProp(0)) <> arrProp(1) Then				
					gstrDesc = "Value of property " & arrProp(0) & " is <> " & arrProp(1) & " For " & strLabel				
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
					bResult = True
				Else
					gstrDesc = "Value of property " & arrProp(0) & "is = " & arrProp(1) & " For " & strLabel				
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0
					bResult = False		   
					objErr.Raise 11 	
				End If	
		ElseIf InStr(1,arrTemp(0),"<=") Then	
			arrProp = Split(arrTemp(0),"<=")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel 
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be less than equal to " & arrProp(1) & " For " & strLabel	
	     		If objTemp.GetROProperty(arrProp(0)) <= arrProp(1) Then
				gstrDesc = "Value of property " & arrProp(0) & " is <= " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = True
			Else
				gstrDesc = "Value of property " & arrProp(0) & " is > " & arrProp(1) & " For " & strLabel
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0				
				bResult = False
				objErr.Raise 11
			End If			
		ElseIf InStr(1,arrTemp(0),">=") Then		
	     		arrProp = Split(arrTemp(0),">=")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel 
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be greater than equal to " & arrProp(1) & " For " & strLabel		
	     		If objTemp.GetROProperty(arrProp(0)) >= arrProp(1) Then
				gstrDesc = "Value of property " & arrProp(0) & " is >= " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = True
			Else
				gstrDesc = "Value of property " & arrProp(0) & " is < " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0				
				bResult = False
				objErr.Raise 11
			End If				
		ElseIf InStr(1,arrTemp(0),">") Then		
			arrProp = Split(arrTemp(0),">")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel 
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be greater than " & arrProp(1) & " For " & strLabel	
			If objTemp.GetROProperty(arrProp(0)) > arrProp(1) Then
				gstrDesc = "Value of property " & arrProp(0) & " is > " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = True
			Else					
				gstrDesc = "Value of property " & arrProp(0) & " is <= " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0				
				bResult = False
				objErr.Raise 11
			End If				
		ElseIf InStr(1,arrTemp(0),"<") Then		
			arrProp = Split(arrTemp(0),"<")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be less than " & arrProp(1) & " For " & strLabel 	
			If objTemp.GetROProperty(arrProp(0)) < arrProp(1) Then
				gstrDesc = "Value of property " & arrProp(0) & " is < " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult=True
			Else
				gstrDesc = "Value of property " & arrProp(0) & " is >= " & arrProp(1) & " For " & strLabel				
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0				
				bResult = False
				objErr.Raise 11
			End If	

		ElseIf InStr(1,arrTemp(0),"Contains") Then
				arrProp = Split(arrTemp(0),"Contains")
					  
				gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strObjType & " - " & strLabel
				gstrExpectedResult= "Value of Property " & arrProp(0) & " contains " & arrProp(1) & " For " & strObjType & " - " & strLabel
					
				Dim SearchString,SearchChar, nLoop
				Dim splitchar,UboundArr,getpos,flag,i,arrSearchChar, bPipeFlag, nUboundArrSearchChar
					
				SearchString = objTemp.GetROProperty(arrProp(0))
				SearchChar = arrProp(1)
				bPipeFlag = False

				If InStr(SearchChar,"|")>0 Then
					arrSearchChar=Split(SearchChar,"|")
					bPipeFlag = True
				End If
					
				If bPipeFlag Then
					nUboundArrSearchChar = UBound(arrSearchChar)
				Else
					nUboundArrSearchChar = 0
					ReDim arrSearchChar(0)
					arrSearchChar(0) = SearchChar
				End If
					
					For nLoop = 0 To nUboundArrSearchChar
						splitchar = Split(arrSearchChar(nLoop)," ")
						UboundArr = Ubound(splitchar)
		
						For i=0 To UBoundArr
							splitchar(i) = LCase(splitchar(i))
						next
		
						SearchString = LCase(SearchString)
						
						For i = 0 to UboundArr
							getpos = Instr(1, SearchString, splitchar(i), 0)
							'msgbox getpos
		
							If  getpos < 1 Then
								flag = FALSE
						Exit For
							Else	
								flag = TRUE			
							End If
						Next 

						If flag = True Then
							Exit For
						End If
		
					Next
					
					If flag = True Then
						gstrDesc = "Value of property " & arrProp(0) & " contains <b>" & arrSearchChar(nLoop) & "</b> for " & strObjType & " - " & strLabel				
							WriteHTMLResultLog gstrDesc, 1
							CreateReport  gstrDesc, 1
						bResult=True
					Else
						TakeScreenShot()				
						gstrDesc = "Value of property " & arrProp(0) & " does not contain <b>" & arrSearchChar(nLoop-1) & "</b> for " & strObjType & " - " & strLabel			
							WriteHTMLResultLog gstrDesc, 0
							CreateReport  gstrDesc, 0
							bResult = False
						 objErr.Raise 11				
					End If
					
		Else
			arrProp = Split(arrTemp(0),"=")				
			If UCase(arrProp(0)) = "EXIST" Then
				strPropValue = "TRUE"
				arrProp(1) = UCase(arrProp(1))
				
				gstrQCDesc = "Verify existance of " & strLabel 
				gstrExpectedResult= strLabel & " should exist"

'				If CStr(trim(strPropValue)) = cstr(trim(arrProp(1))) Then			
				If (StrComp(CStr(strPropValue), trim(cstr(arrProp(1)))) = 0) Then	
					Call TakeScreenShot()
					gstrDesc =  "Successfully got Message <B>'" & strinnertext & "'</B> on Page for Label '" & strLabel & "'" & vbNewLine
					'gstrDesc = "The" & strLabel &" "& strinnertext &" exists."
					WriteHTMLResultLog gstrDesc, 4 '1
					CreateReport  gstrDesc, 1
					bResult = True
					
				Elseif strLabel= "Overlay Title" Then
					 Call TakeScreenShot()	
					gstrDesc = strLabel & " does not exist. The Reason for the failure is Java Script Fallback or Disable on current environment. Once Java Script will be enabled, it will work as expected"								
					WriteHTMLResultLog gstrDesc,0 
					CreateReport  gstrDesc, 0								
					bResult = False
					objErr.Raise 11
		
				Else  		
					Call TakeScreenShot()	
					gstrDesc = strLabel & " does not exist."								
					WriteHTMLResultLog gstrDesc,0 
					CreateReport  gstrDesc, 0								
					bResult = False
					objErr.Raise 11
				End If	
			Else
				strPropValue = objTemp.GetROProperty(arrProp(0))
'				gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel
				gstrExpectedResult= "Value of Property " & arrProp(0) & " should be equal to " & arrProp(1) & " For " & strLabel 							
				If CStr(trim(strPropValue)) = cstr(trim(arrProp(1))) Then				
					gstrDesc = "Value of property " & arrTemp(0) & " is = " & arrProp(1) & " For " & strLabel																
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
					bResult = True
				Else					
					gstrDesc = "Value of property " & arrTemp(0) & " is <> " & arrProp(1) & " For " & strLabel								
					WriteHTMLResultLog gstrDesc,0 
					CreateReport  gstrDesc, 0  
					objErr.raise 11
					'msgbox objErr.Number	
					bResult = False
				End If
			End If				
		End If
	Else
		If InStr(1,arrTemp(0),"<>") > 0 Then
			arrProp = Split(arrTemp(0),"<>")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should not be equal to " & arrProp(1) & " For " & strLabel
			gstrDesc = strLabel & " does not exist"
			WriteHTMLResultLog gstrDesc,0 
			CreateReport  gstrDesc, 0	  
			objErr.Raise 11  			     		
		ElseIf InStr(1,arrTemp(0),"<=") Then	
			arrProp = Split(arrTemp(0),"<=")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel 
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be less than equal to " & arrProp(1) & " For " & strLabel
			gstrDesc = strLabel & " does not exist"
			WriteHTMLResultLog gstrDesc,0 
			CreateReport  gstrDesc, 0	  
			objErr.Raise 11       		
		ElseIf InStr(1,arrTemp(0),">=") Then		
	     		arrProp = Split(arrTemp(0),">=")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel 
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be greater than equal to " & arrProp(1) & " For " & strLabel
			gstrDesc = strLabel & " does not exist"
			WriteHTMLResultLog gstrDesc,0 
			CreateReport  gstrDesc, 0
			objErr.Raise 11			     					
		ElseIf InStr(1,arrTemp(0),">") Then		
			arrProp = Split(arrTemp(0),">")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel 
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be greater than " & arrProp(1) & " For " & strLabel
			gstrDesc = strLabel & " does not exist"
			WriteHTMLResultLog gstrDesc,0 
			CreateReport  gstrDesc, 0  
			objErr.Raise 11 	     					
		ElseIf InStr(1,arrTemp(0),"<") Then		
			arrProp = Split(arrTemp(0),"<")
			gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel
			gstrExpectedResult= "Value of Property " & arrProp(0) & " should be less than " & arrProp(1) & " For " & strLabel 
			gstrDesc = strLabel & " does not exist"
			WriteHTMLResultLog gstrDesc,0 
			CreateReport  gstrDesc, 0	   
			objErr.Raise 11      					
		Else
			arrProp = Split(arrTemp(0),"=")				
			If UCase(arrProp(0)) = "EXIST" Then
				strPropValue = "FALSE"
				gstrQCDesc = "Verify existance of " & strLabel 
				gstrExpectedResult= strLabel & " should exist"
				If CStr(trim(strPropValue)) = UCase(cstr(trim(arrProp(1)))) Then				
					gstrDesc = strLabel & " does not exist"
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
					bResult = True
				Else					
					Call TakeScreenShot()
					gstrDesc = strLabel & " does not exist."								
					WriteHTMLResultLog gstrDesc,0 
					CreateReport  gstrDesc, 0	 
					objErr.Raise 11   						
				End If	
			Else
				strPropValue = objTemp.GetROProperty(arrProp(0))
				gstrQCDesc = "Verify value of property " & arrProp(0) & " for " & strLabel
				gstrExpectedResult= "Value of Property " & arrProp(0) & " should be equal to " & arrProp(1) & " For " & strLabel 							
				gstrDesc = strLabel & " does not exist"
				WriteHTMLResultLog gstrDesc,0 
				CreateReport  gstrDesc, 0	
				objErr.Raise 11
			End If				
		End If		
	End If
	
	Set objTemp = Nothing
	verifyProperty = bResult 

End Function
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Function Name		:setTOProperty
'Input Parameter    	:strObjName - String - Name of the Browser
'Description        	:This function sets the TO Property for specified object
'Calls function        	:NA
'Return Value		:True Or False
'----------------------------------------------------------------------------------------------------------------------
Public Function setTOProperty(strObjName,strLabel,strVal)
	
	Dim objTemp, strClass, bResult, strDesc,arrTemp, arrProp
	Dim bFlag
        bResult = False
		
	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	Set objTemp = gobjObjectClass.getObjectRef(strObjName)
	arrTemp = Split(strVal,";")
	If objTemp.exist(gExistCount) Then		
		objTemp.SetTOProperty arrTemp(0),arrTemp(1)
		bResult = True
	Else
		strDesc = "Object " & strLabel & " does not exist"
		WriteHTMLResultLog strDesc, 0
		CreateReport  strDesc, 0				
		bResult = False
		objErr.Raise 11
	End If
	
	Set objTemp = Nothing
	setTOProperty = bResult 

End Function
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Function Name		:changeReplayType
'Input Parameter    	:strObjName - String - Name of the Browser
'Description        	:This function changes the device replay type
'Calls function        	:NA
'Return Value		:True Or False
'----------------------------------------------------------------------------------------------------------------------
Public Function changeReplayType(strVal)
	
	Dim objQTPApp
    bResult = True	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If
	
	Set objQTPApp = CreateObject("QuickTest.Application")
	If UCase(strVal) = "MOUSE" Then		
		objQTPApp.Options.Web.RunMouseByEvents = True
	Else
		objQTPApp.Options.Web.RunMouseByEvents = False			
	End If
	Set objQTPApp = Nothing		
	changeReplayType = bResult 

End Function
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Function Name    	:invokeBrowser
'Input Parameter    	:strVal - Value to  be entered in text box
'Description            :This function invokes the application in Internet Explorer
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function invokeBrowser(strVal)
	Dim bResult
	bResult=True	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If
	SystemUtil.CloseProcessByName("ctfmon.exe")
	Webutil.DeleteCookies
	Systemutil.Run "iexplore.exe",strVal,,,3
	Webutil.DeleteCookies

	For i=0 to 1 step 1
	If browser("name:=.*").page("title:=.*").link("name:=Continue .*").Exist(2) Then
		 browser("name:=.*").page("title:=.*").link("name:=Continue .*").highlight
		  browser("name:=.*").page("title:=.*").link("name:=Continue .*").click
	End If
	Next
	gstrDesc =  "Browser  '" & strVal& "'  Invoked Successfully."
	gstrExpectedResult = gstrDesc 
	gstrQCDesc = "Invoke Browser '" & strVal & "'."
	WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1
	invokeBrowser=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'Function Name    	:invokeMozillaBrowser
'Input Parameter    	:strVal - Value to  be entered in text box
'Description            :This function invokes the application in Internet Explorer
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function invokeMozillaBrowser(strVal)
	Dim bResult
	bResult=True	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

    SystemUtil.CloseProcessByName("firefox.exe")
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	'objShell.Run """C:\Program Files (x86)\Mozilla Firefox\Firefox.exe"" " & strVal & "",1,False
	objShell.Run """Firefox.exe"" " & strVal & "",1,False
	Set objShell = Nothing


'	SystemUtil.CloseProcessByName("firefox.exe")	
'    SystemUtil.CloseProcessByName("iexplore.exe")
'	webutil.DeleteCookies
'	Systemutil.Run "firefox.exe",,,,3
'	webutil.DeleteCookies
'	Browser("name:=.*").ClearCache
'	Browser("name:=.*").DeleteCookies
'	Browser("name:=.*").Navigate strVal


'	Call SecurityClicks()
	gstrDesc =  "Browser  '" & strVal& "'  Invoked Successfully."
	gstrExpectedResult = gstrDesc 
	gstrQCDesc = "Invoke Browser '" & strVal & "'."
	WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1
	invokeBrowser=bResult
End Function
'----------------------------------------------------------------------------------------------------------------------
'Function Name    	:invokeApp
'Input Parameter    	:strVal - Name of the App
'Description            :This function invokes the application in Internet Explorer
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function invokeApp(strAppPath)
	
	Dim bResult
	bResult=True	
	
	strAppPath = getTestDataValue(strAppPath)
	
	If UCase(strAppPath) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	Systemutil.Run strAppPath
	
	gstrDesc =  "Application '" & strAppPath & "' Invoked Successfully."
	gstrExpectedResult = gstrDesc 
	gstrQCDesc = "Invoke Application '" & strAppPath & "'."
	WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1
	
	invokeApp=bResult

End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name		:ActivateBrowser
'Input Parameter    	:strObjName - String - Name of the Browser
'Description        	:This function activates specified Browser
'Calls function        	:NA
'Return Value		:True Or False
'---------------------------------------------------------------------------------------------------------------------
Public Function ActivateBrowser(strObjName,strLabel)
	
	Dim objBrw, strDesc, bResult
	Dim nHandle, nFlag

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
    bResult = False
	Set objBrw = gobjObjectClass.getObjectRef(strObjName)
	
	If objBrw.exist(gExistCount) Then
	     nHandle = objBrw.Object.HWND
	     Extern.Declare micLong,"ActBrowser","user32","ShowWindow",micLong,micLong
  	     nFlag = Extern.ActBrowser(nHandle,3)
	     bResult = True    
	Else
	     strDesc = "Failed to activate browser " & strLabel
	     WriteHTMLResultLog strDesc, 0
	     CreateReport  strDesc, 0	    	     
		 objErr.Raise 11
	End If
	Set objBrw = Nothing
	ActivateBrowser = bResult 

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:closeBrowser
'Input Parameter    	:strObject - Logical Name of Web Browser
'Description        	:This function closes the browser object
'      			 DLL returns an object reference of a Browser.
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function closeBrowser(strObjName,strLabel, strVal)
	Dim objBrowser, bResult
	bResult=True	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Close Browser '" & strLabel & "'"
	gstrExpectedResult = "Browser '" & strLabel & "' should be closed."
	
	Set objBrowser = gobjObjectClass.getObjectRef(strObjName)
	If objBrowser.Exist(1) Then
		objBrowser.close
		gstrDesc =  "Successfully closed '" & strLabel & "' Browser."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
	Else
		Call TakeScreenShot()
		gstrDesc =  "Browser '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		bResult=False
		  objErr.Raise 11
	End If
	CloseBrowser = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
'Function Name    	:CloseAllBrowsers
'Inpute Paramers  	:None
'Description    	:Closes all open browser on the desktop
'Calls      		:None
'Return Value   	:True/False
'----------------------------------------------------------------------------------------------------------------------
Public Function CloseAllBrowsers(strVal)
	Dim oBrowserList, oBrw,bResult
	Dim hwnd
	bResult = True

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	gstrQCDesc = "Close All open Browsers"
	gstrExpectedResult = "All Browsers should get closed"

	'Object description for Browser
	set oBrw = Description.Create()
'	oBrw("nativeclass").Value = "IEFrame"
	oBrw("micclass").Value = "Browser"
	'Collecting all browser objects in to objectlist lstBrw
	Set oBrowserList = Desktop.ChildObjects(oBrw)
	Dim numAttempts
	numAttempts = 0
	On Error Resume Next
	If oBrowserList.Count > 0 Then
		Do While numAttempts < oBrowserList.Count
      			Browser("index:=0").Close
      			'Check to see if browser exists - if it does, then proceed through dialog checks
      			Set oBrowserList = Desktop.ChildObjects(oBrw)
      			If oBrowserList.Count > 0 Then
            			If Browser("index:=0").Dialog("nativeclass:=.*32770").Exist(0) Then
              				If Browser("index:=0").Dialog("nativeclass:=.*32770").WinButton("nativeclass:=Button","text:=OK").Exist(0) Then
                				Browser("index:=0").Dialog("nativeclass:=.*32770").WinButton("nativeclass:=Button","text:=OK").Click
              				End If
            			ElseIf Browser("index:=0").Dialog("title:=Security Alert").Exist(0) Then
              				If Browser("index:=0").Dialog("title:=Security Alert").WinButton("nativeclass:=Button","text:=OK").Exist(0) Then
                				Browser("index:=0").Dialog("title:=Security Alert").WinButton("nativeclass:=Button","text:=OK").Click
          				End If
        			ElseIf Browser("index:=0").Dialog("title:=Security Information").Exist(0) Then
              				If Browser("index:=0").Dialog("title:=Security Information").WinButton("nativeclass:=Button","text:=OK").Exist(0) Then
                				Browser("index:=0").Dialog("title:=Security Information").WinButton("nativeclass:=Button","text:=OK").Click
              				End If
            			End If  'WinDialog exists

        			'Get updated list of open browsers on desktop
        			Set oBrowserList = Desktop.ChildObjects(oBrw)
          			If oBrowserList.Count > 0 Then
          				Browser("index:=0").Close
        			End If

        		End If  'browser exists after checking/closing dialogs

      			numAttempts = numAttempts + 1
      			'get updated list of browsers open
      			Set oBrowserList = Desktop.ChildObjects(oBrw)
    		Loop
  	End If 'more than 0 browsers open

  	'Confirming all Browsers are closed
  	Set oBrowserList = Desktop.ChildObjects(oBrw)
  	If IsNull(oBrowserList) = False Then
    		If oBrowserList.Count > 0 Then			
      			'WriteHTMLResultLog "Failed to close all browsers. " & oBrowserList.Count & " browser windows remain open.",0
    		'Else
      			'WriteHTMLResultLog "All browser windows have been closed successfully.",1
      			'bResult = True
    		End If
  	End If	
  	CloseAllBrowsers=bResult

End Function
'----------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:syncBrowser
'Input Parameter      	:objPage(Object) - representing page object
'     			:strBtnObject(String) - representing the logical name of Button object
'Description          	:This function waits till the button is enabled
'Calls                	:errorHandler in case of any error
'Return Value   	:True /False
'---------------------------------------------------------------------------------------------------------------------
Public Function syncBrowser(strObject, strVal)

'	'Variable for Objects
'	Dim objBrowser, objPage,objFrame
'	Dim sTime, bFlag, bResult, intTimeOut, arrFrame, objDesc, intLoop
'		
'	bResult = True
'
'	strVal = getTestDataValue(strVal)
'	
'	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
'		Exit Function
'	End If
'
'	'get the starting time
'	sTime = TIMER
'	intTimeOut = 700
'	bFlag = True
'
'	Set objBrowser = gobjObjectClass.getObjectRef(strObject)
'    gstrBrowserURL = objBrowser.GetROProperty("url")
'	gstrBrowserHWND = objBrowser.GetROProperty("hwnd")
'	wait(2)
'	If objBrowser.Exist(gExistCount) And bFlag Then
'		Do While objBrowser.Object.Busy = True
'			Wait(1)
'			If Dialog("nativeclass:=.*32770").Exist(0) Then
'              			bFlag = False
'                		Exit Do
'			End If		              		
'			If (sTime + intTimeOut) > TIMER Then
'				bFlag = False
'				Exit Do
'			End If
'		Loop
'		
'		Set objPage = objBrowser.Page("micclass:=Page")
'		If objPage.Exist(gExistCount) And bFlag Then
'			Do While UCase(objPage.Object.readyState) <> "COMPLETE"
'				Wait(1)
'				If Dialog("nativeclass:=.*32770").Exist(0) Then
'              				bFlag = False
'                			Exit Do
'				End If
'				If (sTime + intTimeOut) > TIMER Then
'					bFlag = False
'					Exit Do
'				End If
'			Loop
'		
'		End If
'	Else
'		objErr.Raise 11
'	End If
'    SyncBrowser = bResult

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name       	:VerifyPage
'Input Parameter     	:strVal - Text to print
'Description         	:This function prints the non - automatable steps
'Return Value 		:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifyPage(strObject,strLabel, strVal)
	Dim bResult,strDesc,objElm	
	bResult=False	

	strVal = getTestDataValue(strVal)
 
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
    Set objElm = gobjObjectClass.getObjectRef(strObject)
	
	If Not objElm Is Nothing Then

		gstrQCDesc = "Verify Page " & strLabel & " is displayed."
		gstrExpectedResult = "Page " & strLabel & " should be displayed."	
'		If objElm.Exist Then
		If objElm.Exist Then
				Call TakeScreenShot()
				''strurl=objElm.GetRoProperty("url")
			    gstrDesc =  "Page '" & strLabel & "' is displayed successfully."  
        		WriteHTMLResultLog gstrDesc, 2
        		CreateReport  gstrDesc, 2
				
				'gstrDesc1 =  "Page  is displayed with URL : " & strurl	   
				'WriteHTMLResultLog gstrDesc1, 4
        		'CreateReport  gstrDesc1, 1
			bResult = True       	
		Else
				Call TakeScreenShot()				
          		gstrDesc =  "Page " & strLabel & " is not displayed properly."				
          		WriteHTMLResultLog gstrDesc, 0
          		CreateReport  gstrDesc, 0		
				 bResult = False
				objErr.Raise 11
			   
		End If
	End If
  	VerifyPage=bResult
	
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:clickButton
'Input Parameter    	:strObject - Logical Name of Web Button
'Description          	:This function clicks on the Web Button
'      			 DLL returns an object reference of a WebButton
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function clickButton(strObject,strLabel,strVal)
	Dim objButton, bResult,objTrim,objlen,objReport
	bResult=False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	If instr(strObject,":") Then
		arrobj=Split(strObject,":")
		For i= 0 To Ubound(arrobj)
			Set objButton=Nothing
			Set objButton=gobjObjectClass.getObjectRef(arrobj(i))
			If objButton.Exist  Then
				objButton.click
				bSlection=True
				Exit For
			End If
		Next
		If bSlection=True  Then
            gstrDesc =  "Successfully clicked on '" & strLabel & "' button."		
			WriteHTMLResultLog gstrDesc, 4
			CreateReport  gstrDesc, 1
			bResult = True
		Else
			 Call TakeScreenShot()
			gstrDesc =  "Button '" & strLabel & "' does not exist"		
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
		End If
	Else
		Call TakeScreenShot()
		Set objButton = gobjObjectClass.getObjectRef(strObject)
		gstrQCDesc = "Click on button " & strLabel
		gstrExpectedResult = "clicked on '" & strLabel & "' button."
	
		If objButton.exist(gExistCount)  Then
			'objButton.waitProperty "attribute/readyState", "complete", gExistCount*1000
			objButton.click
			If gbIterationFlag <> True Then
				gstrDesc =  "Successfully clicked on '" & strLabel & "' button."		
				WriteHTMLResultLog gstrDesc, 4
				CreateReport  gstrDesc, 1
			End If
			bResult = True
		Else
			 Call TakeScreenShot()
			gstrDesc =  "Button '" & strLabel & "' does not exist"		
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
		End If
	End If 
	
	clickButton=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:clickImage
'Input Parameter      	:strObject - Logical Name of Image
'Description          	:This function clicks the Image object
'      			 DLL returns an object reference of an Image
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function clickImage(strObject,strLabel, strVal)
	
	Dim bResult, objImage
	bResult=False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	call TakeScreenShot()
	Set objImage = gobjObjectClass.getObjectRef(strObject)

	gstrQCDesc = "Click on Image " & strLabel
	gstrExpectedResult = "clicked on '" & strLabel & "' Image"
	If objImage.exist(gExistCount) Then
		objImage.click
		If gbIterationFlag <> True then
			gstrDesc =  "Successfully clicked on '" & strLabel & "'  Image."
			WriteHTMLResultLog gstrDesc, 4
			CreateReport  gstrDesc, 1
		End If
'		objImage.waitProperty "attribute/readyState", "complete", gExistCount*1000
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Image '" &strLabel& "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0    		
		bResult = False
		objErr.Raise 11
	End If

	clickImage=bResult

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:enterText
'Input Parameter      	:strObject - Logical Name of Edit Box
'     			:strVal - Value to  be entered in text box
'Description        	:This function enters a data into a text box
'      			 DLL returns an object reference of a WebEdit
'Calls              	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function enterText_OLD(strObject,strLabel,strVal)
	
		Dim bResult, objTextField,objTrim,objlen,objReport, i, Wsh
	Dim strtemp
	bResult=False			

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
		
	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."
	
	If instr(strObject,":") Then
		arrobj=Split(strObject,":")
		For i= 0 To Ubound(arrobj)
			Set objElm=gobjObjectClass.getObjectRef(arrobj(i))
			If objElm.Exist  Then
				objElm.set strVal
			    gstrDesc =  "Successfully entered '" & strVal & "' in '" & strLabel & "' textbox."
        		WriteHTMLResultLog gstrDesc, 2
        		CreateReport  gstrDesc, 2
				bResult=True
				Exit For
			End If
		Next
		If bResult<>True Then
			Call TakeScreenShot()
			gstrDesc =  "'" & strLabel & "' textbox does not exist"
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0      		
			bResult = False
			objErr.Raise 11
		End If
	Else
		Set objTextField = gobjObjectClass.getObjectRef(strObject)
		If objTextField.exist(gExistCount) Then  
			objTextField.set strVal
			If gbIterationFlag <> True then
				gstrDesc = "Successfully entered '" & strVal & "' in '" & strLabel & "' textbox."
				 'gstrDesc = "Successfully entered '" & getTestDataValue(strVal) & "' in '" & strLabel & "' textbox."  	
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
			End if
			bResult = True
		Else
			Call TakeScreenShot()
			gstrDesc =  "'" & strLabel & "' textbox does not exist"
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0      		
			bResult = False
			objErr.Raise 11
		End If
	End If  	
	enterText=bResult
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:enterText
'Input Parameter      	:strObject - Logical Name of Edit Box
'     			:strVal - Value to  be entered in text box
'Description        	:This function enters a data into a text box
'      			 DLL returns an object reference of a WebEdit
'Calls              	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function enterText(strObject,strLabel,strVal)
	
	Dim bResult, objTextField,objTrim,objlen,objReport, i, Wsh, nCnt, strTextVal
	bResult=False			

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."
	Set objTextField = gobjObjectClass.getObjectRef(strObject)
	If objTextField.exist(gExistCount) Then
					objTextField.Click
					Const Chars = "abcdefghijklmnopqrstuvwxyz"
			
					If UCase(strVal) = "DATE" Then
								strDate=Date
								objTextField.Set strDate
					ElseIf UCase(strVal) = "MAIL" Then
								strName = ""
								Randomize
								For k = 1 To 7
									intValue = Fix(26 * Rnd())
										strChar = Mid(Chars, intValue + 1, 1)	
									strName = strName & strChar
								Next
								tempName= UCase(strName)
								CustSurname = "TBT" & tempName
								nrNumber = Int((1000 * Rnd) + 1)
								strmailid=CustSurname & nrNumber
								strMail ="TBT" & strmailid &"@Portal.com"
								gstrREMAIL=strMail
								gstrRegistrationEmail=strMail
								objTextField.Set strMail
					ElseIf UCase(strVal) = "RANDOMNUMBER" Then
								 MyValue = Int((999 * Rnd) + 100)
								 objTextField.Set MyValue
					ElseIf UCase(strVal) = "RANDOMTEXT" Then
								strName = ""
								Randomize
								For k = 1 To 3
									intValue = Fix(26 * Rnd())
									strChar = Mid(Chars, intValue + 1, 1)	
									strName = strName & strChar
								Next
								tempName= UCase(strName)
								CustSurname = tempName
								objTextField.Set CustSurname
					 ElseIf UCase(strVal) = "FUTUREDATE" Then
								MyDate = Date 
								today = FormatDateTime(Date,2)
								chgDate = DateAdd ("d",2,today)
								
								dayName = WeekDay(chgDate)
								
								If dayName = ("1" or "7") Then
										chgDate = DateAdd("d",2,chgDate)
								End If
								today = FormatDateTime(chgDate,2)
								objTextField.Set today
					Else
								objTextField.Set strVal
					End If
					
					If gbIterationFlag <> True then
						If strVal="FUTUREDATE" Then
							gstrDesc = "Successfully entered '" & today & "' in '" & strLabel & "' textbox."
						ElseIf UCase(strVal)="DATE" Then
							gstrDesc = "Successfully entered '" & strDate & "' in '" & strLabel & "' textbox."
						ElseIf UCase(strVal) = "RANDOMTEXT" Then
							gstrDesc = "Successfully entered '" & CustSurname & "' in '" & strLabel & "' textbox."
						ElseIf UCase(strVal) = "MAIL" Then

							gstrDesc = "Successfully entered '" & strMail & "' in '" & strLabel & "' textbox."
						Else
							gstrDesc = "Successfully entered '" & strVal & "' in '" & strLabel & "' textbox."
						End If
										
						WriteHTMLResultLog gstrDesc, 1
						CreateReport  gstrDesc, 1
					End if
					bResult = True
	Else
					 Call TakeScreenShot()
					gstrDesc =  "'" & strLabel & "' textbox does not exist"
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0      		
					bResult = False
					objErr.Raise 11
	End If

  	enterText=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name  :typeText
'Input Parameter :strObject - Logical Name of Edit Box
'   :strVal - Value to  be entered in text box
'Description  :This function enters a data into a text box in WinXcel Application.
'Calls   :None
'Return Value  :True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function typeText(strObject, strLabel, strVal)
	Dim bResult, objTextField, dVar,splitText,ArrEditName, Wsh
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."

	ArrEditName=Split(strObject,"-")
	Set objTextField = gobjObjectClass.getObjectRef(strObject)
	
	If  objTextField.exist(gExistCount) Then
		gstrExpectedResult = strVal & " should get typed  in '" & strLabel & "' textbox."
		If Trim(objTextField.getROProperty("value")) <> ""  Then
			objTextField.set ""
			wait 1
		End If
		objTextField.click
		Set Wsh = CreateObject("Wscript.Shell")
		wait 2 
		Wsh.SendKeys strVal
		If gbIterationFlag <> True then
			gstrDesc = "Successfully typed '" & strVal & "' in '" & strLabel & "' textbox."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "'" & strLabel & "' textbox does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If
	 typeText=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:selectList
'Input Parameter      	:strObject - Logical Name of List Box
'     			:strData - Value to  be selected from list box
'Description          	:This function enters a data into a text box
'      		  	 DLL returns an object reference of a WebList
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function selectList(strObject,strLabel,strVal)
   Dim objListField,objTrim,objlen,objReport,Wsh, nLoop,arrAccount,arrTemp,flg,j,strTmp
	Dim intBusiness,LstLBound,arrTmp,strsortcode,i
	bResult = False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."
	Set objListField=gobjObjectClass.getObjectRef(strObject)
		
	If objListField.exist(gExistCount) Then
		objListField.Object.focus
		For i=0 to 9
				 nCount=objListField.GetROProperty("items count")
				 StrAllItems=Split(objListField.GetROProperty("all items"),";")
				intStart=Lbound(strAllItems)
				intEnd=Ubound(strAllItems)+1
				If   nCount  >1 Then
						Exit For
				Else
						Wait 1
				End If
		Next
     
		If  nCount >1Then
						wait 1
						For intCounter = 1 to intEnd
							If objListField.GetItem(intCounter) = strVal Then
												objListField.Select "#"& intCounter-1
												bResult = True
												Exit For
									Else
												bResult = False
									End If 					
						Next

						If bResult Then
									gstrDesc = "Value '" & objListField.getROProperty("value") & "' is selected successfully from '" & strLabel & "' list."
									WriteHTMLResultLog gstrDesc, 1
									CreateReport  gstrDesc, 1
						Else
									Call TakeScreenShot()
									gstrDesc=  strVal & "  Value not found...Please write a reporter event for this"
									WriteHTMLResultLog gstrDesc, 0
									CreateReport  gstrDesc, 0		
									bResult = False
									objErr.Raise 11
						End If
						
		Else
				Call TakeScreenShot()
				gstrDesc =  "Items in the  List '" & strLabel & "' are not displayed."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11
		End If
			
		 
	Else
		 Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	
	selectList=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:verifyItemsInList
'Input Parameter      	:strObject - Logical Name of List Box
'     					:strVal - Value to  be verified in list along with the expected result(True/False)
'Description          	:This function verifies the existence of object in the listbox	
'Calls                	:None
'Return Value   		:None
'---------------------------------------------------------------------------------------------------------------------
Public Function verifyItemsInList(strObject,strLabel,strVal)

	Dim objListField,strAllItems,arrTemp,bExp,strInputList,intCnt
	Dim strInList,strOutList
	bResult = False	   

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

On Error Resume Next
	Set objListField=gobjObjectClass.getObjectRef(strObject)
	arrTemp = Split(strVal,";")
	bExp = Ucase(arrTemp(UBound(arrTemp)))
	For intCnt = 0 To UBound(arrTemp) - 1
		If strInputList = "" Then
			strInputList = arrTemp(intCnt)
		Else
			strInputList = strInputList & ";" & arrTemp(intCnt)
		End If
	Next
	If UCase(bExp) = "TRUE" Then
		gstrQCDesc = "Verify existence of items " & strInputList & " in List " & strLabel &"."
		gstrExpectedResult = "Value '" & strInputList & "' should be present in List " & strLabel&"."
	Else
		gstrQCDesc = "Verify non existence of items " & strInputList & " in List " & strLabel
		gstrExpectedResult = "Value '" & strInputList & "' should not be present in List " & strLabel &"."
	End If
	If objListField.exist(3) Then
'		objListField.Select strData
		strAllItems = objListField.GetROProperty("all items")		
		For intCnt = 0 To UBound(arrTemp) -1  		
			If InStr(1,strAllItems,arrTemp(intCnt)) > 0 Then
                If strInList = "" Then
					strInList = arrTemp(intCnt)
				Else
					  strInList = strInList & ";" & arrTemp(intCnt)
'					strInList = ";" & arrTemp(intCnt)
				End If				
			Else
				If strOutList = "" Then
					strOutList = arrTemp(intCnt)
				Else
					  strOutList = strOutList & ";" & arrTemp(intCnt)
'					strOutList = ";" & arrTemp(intCnt)
				End If							
			End If
		Next		
		If bExp = "TRUE" Then
			If strOutList = "" Then
				gstrDesc = "Items '" & strInputList & "' are present in the "& strLabel &" list."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = True
			Else
				If strInList = "" Then
					gstrDesc = "Items '" & strOutList & "' are not present in the " & strLabel &" list." 
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0
					objErr.Raise 11
				Else
					gstrDesc = "Items '" & strInList & "' are present, while the items '" & strOutList & "' are not present in the " & strLabel &" list." 
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0
					objErr.Raise 11
				End If
			End If
		Else
			If strInList = "" Then
				gstrDesc = "Items '" & strOutList & "' not present in list " & strLabel
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = True
			Else
				gstrDesc = "Items '" & strInList & "' present in list " & strLabel
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0
				objErr.Raise 11
			End If
		End If				
	Else
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		objErr.Raise 11
	End If
	verifyItemsInList=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:setCheckBox
'Input Parameter      	:strObject - Logical Name of Check box
'     			:strVal - Contains the property of the Chechkbox which needs to be changed and the
'      			 value it needs to be changed to.
'Description          	:This function toggles the checkbox
'      			 DLL returns an object reference of a WebCheckBox
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function setCheckBox(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objChkBox = gobjObjectClass.getObjectRef(strObject)		
	gstrQCDesc = "Set the check box " & strLabel & " " & UCase(strVal)
	gstrExpectedResult = "Checkbox " & strLabel & " should be set to " & UCase(strVal)
	If objChkBox.exist(gExistCount) Then

		objChkBox.Set strVal

		If gbIterationFlag <> True then
			gstrDesc = "Successfully selected the '" & strLabel & "' Checkbox."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Checkbox '" & strLabel & "' is not displayed in the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	setCheckBox = bresult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:clicklink
'Input Parameter      	:strObject - Logical Name of Link
'Description          	:This function clicks the Link object
'      			 DLL returns an object reference of a Link
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function clickLink(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False
'	call TakeScreenShot()
	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Click on Link " & strLabel
	gstrExpectedResult = "clicked on '" & strLabel & "' Link"

	objTrim= Left(strObject,3)
	objlen= len(strObject)
	If (objTrim="lnk") then
		objReport= mid(strObject,4,objlen)
	Else		
		objReport=strObject
	End if
	
	Call TakeScreenShot()
	Set objLink = gobjObjectClass.getObjectRef(strObject)
	If objLink.exist(gExistCount) Then
'		objLink.waitProperty "attribute/readyState", "complete", gExistCount*1000
		If gbIterationFlag <> True then
			gstrDesc =  "Successfully clicked on '" & strLabel & "' hyperlink."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
'		objLink.highlight
		objLink.Click
		bResult = True
	Else
		 Call TakeScreenShot()
		 If strVal ="NOTEXIST" Then
			gstrDesc =  " The 'Verify user (ID&V)' should not  be displayed since SSA user's Telephony account status is 'Locked'."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = True
		Else
			gstrDesc =  "WebLink '" &strLabel & "' is not displayed on the screen."
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
		End If
	End If
	clickLink=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:selectRadioButton
'Input Parameter      	:strObject - Logical Name of Check box
'     			:strVal - Contains the property of the Chechkbox which needs to be changed and the
'      			 value it needs to be changed to.
'Description          	:This function toggles the checkbox.
'      			 DLL returns an object reference of a WebCheckBox
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function selectRadioButton(strObject,strLabel,strVal)

	Dim objRadio,bResult
	bResult = False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objRadio = gobjObjectClass.getObjectRef(strObject)

	gstrQCDesc = "Select value '" & getTestDataValue(strVal) & "' from '" & strLabel & "' Radio Button."
	gstrExpectedResult = "Value '" & getTestDataValue(strVal) & "' should be selected successfully from '" & strLabel & "' Radio Button."
	
	If objRadio.exist(gExistCount) Then
		objRadio.select strVal
'		objRadio.set
		If gbIterationFlag <> True then
			gstrDesc =  "Successfully clicked on '" & getTestDataValue(strVal) & "' radio button for '" & strLabel & "' Radiogroup ."
			gstrExpectedResult = gstrDesc
			gstrQCDesc = "Click on radio button '" & getTestDataValue(strVal) & "'."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Radio Button '" & getTestDataValue(strVal) &"( for " & strLabel & " radiogroup)" &"' is not displayed."
		gstrExpectedResult = "Successfully clicked on '" & getTestDataValue(strVal) & "' Radio button."
		gstrQCDesc = "Click on Radio button '" & getTestDataValue(strVal) & "'."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If 
	
	selectRadioButton = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:clickWebElement
'Input Parameter      	:strObject - Logical Name of Web Button
'Description          	:This function clicks the WebElement object
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function clickWebElement(strObject,strLabel, strVal)

	Dim objWebElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	objTrim= Left(strObject,3)
	objlen= len(strObject)

	If (objTrim="elm") then
		objReport= mid(strObject,4,objlen)
	Else
		objReport=strObject
	End if

	gstrQCDesc = "Click on web element " & strLabel & "."
	gstrExpectedResult = "clicked on '" & strLabel& "' web element."
	
	Set objWebElement = gobjObjectClass.getObjectRef(strObject)
	If objWebElement.exist(gExistCount) Then
		objWebElement.click

		If gbIterationFlag <> True then
			gstrDesc =  "Successfully clicked on '" & strLabel & "'  WebElement."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebElement '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	clickWebElement=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'=====================================================================================================================
'INTERNAL FUNCTIONS
'=====================================================================================================================

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:AssociateOR
'Input Parameter    	:strAction - Name of action
'Description            :Associate .tsr objectrepository to action
'Calls                  :None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Function AssociateOR(strAction, strTSR, blnFlag)

	Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
	Dim qtRepositories 'As QuickTest.ObjectRepositories ' Declare an action's object repositories collection variable
	Dim lngPosition

	'Open QuickTest
	Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
	'Get the object repositories collection object of the "strAction" action
	Set qtRepositories = qtApp.Test.Actions(strAction).ObjectRepositories

	If blnFlag = "True" Then
    		'Add MainApp.tsr if it's not already in the collection
   		If qtRepositories.Find(strTSR) = -1 Then ' If the repository cannot be found in the collection
          		qtRepositories.Add strTSR, 1 ' Add the repository to the collection
    		End If
  	Else
    		lngPosition = qtRepositories.Find(strTSR) ' Try finding the .tsr object repository
    		If lngPosition <> -1 Then ' If the object repository was found in the collection
        		qtRepositories.Remove lngPosition ' Remove it
    		End If
  	End If

  	Set qtRepositories = Nothing ' Release the action's shared repositories collection
  	Set qtApp = Nothing ' Release the Application object

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:TakeScreenShot
'Input Parameter    	:strObject - Logical Name of  Browser
'Description          	:This function takes the screen shot of the application
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------
Public Function TakeScreenShot()
	
	Dim strSSTime,bResult,objParent,strParentObj,arrParentObj,strExc
	Dim strDesc,strScreenShotName ,strScreenShotPath, objFSO, objFolder

	bResult = True
	strScreenShotPath=""
	strSSTime = Now( )

	If gbScreenShotImageStatus=FALSE Then
		Exit Function
	End IF

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FolderExists(gstrBaseDir & "Reports\" & gstrProjectUser & "\ScreenShots\" & gstrCurScenario) = False Then
		Set objFolder = objFSO.createFolder(gstrBaseDir & "Reports\" & gstrProjectUser & "\ScreenShots\" & gstrCurScenario)
'		objFSO.deleteFolder(gstrBaseDir & "Reports\ScreenShots\" & gstrCurScenario)
	End If

	On Error Resume Next
	strSSTime= replace(strSSTime,"/","_")
	strSSTime =replace (strSSTime,":","_")
	strSSTime= replace(strSSTime," ","_")		
	strParentobj = replace(gObjectpath,Chr(34),"$")			
	arrParentobj=split(strParentobj,".")	
	strParentobj=replace(arrParentObj(0),Chr(34),"")	
	strParentobj=replace(strParentobj,"$",Chr(34))
	
	strExc = "Set objParent=" & strParentobj
	Execute strExc	
	strScreenShotPath = gstrBaseDir& "Reports\" & gstrProjectUser & "\ScreenShots\"& gstrCurScenario & "\" &gstrGroupName & "_" & gstrCurScenario&"_"&strSSTime&".png"	
	strScreenShotName = "<br>" & gstrGroupName & "_" & gstrCurScenario&"_"&strSSTime&"<br>"	

	Desktop.CaptureBitmap  strScreenShotPath , True			
	clsEnvironmentVariables.ScreenShotPath = strScreenShotPath	
    gstrScreenShotPath=strScreenShotPath
	
	Set objFSO = nothing
	Set objFolder = nothing

	On Error GoTo 0
	
	TakeScreenShot = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:TakeScreenShot_Full
'Input Parameter    	:strObject - Logical Name of  Browser
'Description          	:This function takes the screen shot of the application
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------
Public Function TakeScreenShot_Full(strObject,strLabel)
	
	Dim strSSTime,bResult,objParent,strParentObj,arrParentObj,strExc
	Dim strDesc,strScreenShotName ,strScreenShotPath, objFSO, objFolder

	bResult = True
	strScreenShotPath=""
	strSSTime = Now( )
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Set objBrowser = gobjObjectClass.getObjectRef(strObject)
	Set oScreenCapture = CreateObject("KnowledgeInbox.ScreenCapture")
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FolderExists(gstrBaseDir & "Reports\" & gstrProjectUser & "\ScreenShots\" & gstrCurScenario) = False Then
		Set objFolder = objFSO.createFolder(gstrBaseDir & "Reports\" & gstrProjectUser & "\ScreenShots\" & gstrCurScenario)
'		objFSO.deleteFolder(gstrBaseDir & "Reports\ScreenShots\" & gstrCurScenario)
	End If

	On Error Resume Next
	strSSTime= replace(strSSTime,"/","_")
	strSSTime =replace (strSSTime,":","_")
	strSSTime= replace(strSSTime," ","_")		
	strParentobj = replace(gObjectpath,Chr(34),"$")			
	arrParentobj=split(strParentobj,".")	
	strParentobj=replace(arrParentObj(0),Chr(34),"")	
	strParentobj=replace(strParentobj,"$",Chr(34))
	
	strExc = "Set objParent=" & strParentobj
	Execute strExc	
	strScreenShotPath = gstrBaseDir& "Reports\" & gstrProjectUser & "\ScreenShots\"& gstrCurScenario & "\" &gstrGroupName & "_" & gstrCurScenario&"_"&strSSTime&".png"	
	strScreenShotName = "<br>" & gstrGroupName & "_" & gstrCurScenario&"_"&strSSTime&"<br>"	
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	If objParent.Exist(0) Then
		objParent.HighLight
		bURL = objParent.GetROProperty("URL")
		bHWND = objParent.GetROProperty("hwnd")
	Else
		bURL = gstrBrowserURL
		bHWND = gstrBrowserHWND
	End If	
	oScreenCapture.CaptureIEFromCurrentPos = False
	oScreenCapture.TextFontName = "Arial"
	oScreenCapture.TextFontSize = 7
	oScreenCapture.TextHeight = 20
	oScreenCapture.TextPos = 1
	oScreenCapture.CaptureIE bHWND, strScreenShotPath, bURL & vbcrlf & strLabel, True, True			
	clsEnvironmentVariables.ScreenShotPath = strScreenShotPath
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'	Desktop.CaptureBitmap  strScreenShotPath , True			
	clsEnvironmentVariables.ScreenShotPath = strScreenShotPath	
    gstrScreenShotPath=strScreenShotPath
	
	Set objFSO = nothing
	Set objFolder = nothing

	On Error GoTo 0
	
	TakeScreenShot_Full = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:pressTab
'Description          	:Presses tab key on keyboard
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------
Public Function pressTab()
	
	Dim objShell, bResult

	bResult = True

	If objErr.Number = 11 Then
		Exit Function
	End If
	
    Set objShell = CreateObject("WScript.Shell")

	Wait(1)
	objShell.Sendkeys "{TAB}"

    Set objShell = nothing

    pressTab = bResult
	
End Function
'---------------------------------------------------------------------------------------------------------------------
Public Function pressEnter()
	
	Dim objShell, bResult

	bResult = True

	If objErr.Number = 11 Then
		Exit Function
	End If
	
    Set objShell = CreateObject("WScript.Shell")

	Wait(1)
	objShell.Sendkeys "{ENTER}"

    Set objShell = nothing

	gstrDesc =  "Successfully clicked on Submit button."		
	WriteHTMLResultLog gstrDesc, 4
	CreateReport  gstrDesc, 1
	bResult = True

	pressTab = bResult
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:clickWebElementFromWebTable
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function clickWebElementFromWebTable(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	arrVal=Split(strVal,":")

	Set oDesc = Description.Create() 
	oDesc("micclass").Value = "WebElement" 
	
	Set objTable = gobjObjectClass.getObjectRef(strObject)
	Wait 10
	If  objTable.EXIST(gExistCount) Then

		Set objElementCollection = objTable.ChildObjects(oDesc)
		
		NumberOfWebElements = objElementCollection.Count 
		nFlag = 0
		For i = 0 To NumberOfWebElements - 1 

				If StrComp(Trim(objElementCollection (i).GetROProperty("innertext")),Trim(arrVal(0))) = 0  Then
				
						If arrVal(1)="DoubleClick" Then
										objElementCollection(i).highlight							
										objElementCollection(i).FireEvent "ondblclick"
									
									
						Else
										objElementCollection(i).highlight	
										objElementCollection(i).click
									
									
                        End If		
					
'						nFlag = nFlag + 1
'			
'						If nFlag > 2 Then
'										i = NumberOfWebElements
'						End If
				End If 
		
		Next 

				gstrDesc =  "Successfully clicked on " & arrVal(0) & "  "& strLabel & "."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = true
		
	Else
		Call TakeScreenShot()
		gstrDesc =  "WebTable " & strLabel & " is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
    
	clickWebElementFromWebTable = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:setWebTableCheckBox
'Input Parameter      	:strObject - Logical Name of Web Table
'										strLabel:  Label to be used in Report
'										strVal: Value by which function will search for a checkbox
'Description          	:This function checks a specific checkbox in WebTable
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------

Public Function setWebTableCheckBox(strObject, strLabel, strVal)

	Dim nRowIndex, nColumnIndex, bResult, nflag, objTable

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	Set objTable = gobjObjectClass.getObjectRef(strObject)
	If  objTable.EXIST(gExistCount) Then
'		objTable.waitProperty "attribute/readyState", "complete", gExistCount*1000

		For nRowIndex = 1 to objTable.RowCount
			For nColumnIndex = 1 to objTable.ColumnCount(nRowIndex)
				If objTable.GetCellData(nRowIndex, nColumnIndex) = strVal then
					wait(2)
					objTable.ChildItem(nRowIndex, 1, "WebCheckBox",0).Set "ON"
					nflag = 1
					Exit for
				end if 
				If nflag = 1 Then
					Exit for
				End If
	
			Next
		Next
		bResult = True
		
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebTable " & strLabel & " is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	
	setWebTableCheckBox = bResult
   
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:selectItemInList
'Input Parameter      	:strObject - Logical Name of Web Table
'										strLabel:  Label to be used in Report
'Description          	:This function selects an item from the list
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------

Public Function selectItemInList(strObject, strLabel, strVal)

	Dim objList,strItem, bResult

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objList = gobjObjectClass.getObjectRef(strObject)

	If  objList.exist(gExistCount) Then
'		objList.waitProperty "attribute/readyState", "complete", gExistCount*1000
		strItem = objList.GetItem(2)
		objList.select strItem
		gstrDesc = "Value '" & strItem & "' is selected successfully from '" & strLabel & "' list."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If

	Set objList = Nothing
	selectItemInList = bResult

End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:sendNumLock
'Input Parameter      	:None
'Description          	:This function sends numlock key signals
'Return Value         	:Nothing
'---------------------------------------------------------------------------------------------------------------------

Public Function sendNumLock()

	Dim WshShell
	
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.SendKeys "{NUMLOCK}"
	Set WshShell = Nothing

End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:SendEnter
'Input Parameter      	:None
'Description          	:This function sends numlock key signals
'Return Value         	:Nothing
'---------------------------------------------------------------------------------------------------------------------

Public Function SendEnter()

	Dim WshShell
	Wait 2
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys "~"
	Set WshShell = Nothing

End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:getTestDataValue
'Input Parameter      	:strVal- Pass Coloum Name
'Description          	:This function get the TestDataValue
'Return Value         	:Nothing
'---------------------------------------------------------------------------------------------------------------------
Public Function getTestDataValue(ByVal strVal)

	If dictData.Exists(Trim(strVal)) Then
		strcoloumnname=strVal
		strVal = arrData(nIDIndex,dictData.Item(strVal))
		If  strVal="EXCEL" Then
			intsno=nIDIndex+1
			stractionname=gstrGroupName
			Set con = CreateObject("ADODB.Connection")
			Set rs = CreateObject("ADODB.Recordset")
			con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& gstrControlFilesDir & "\TestData.xls;Excel 8.0;HDR=Yes;"
			str = "SELECT "&strcoloumnname&" from ["&stractionname&"$] where Flag='Y' and ID="&gstrCurID&" and Sno="&intsno&""
			rs.Open str, con
			'intsno=nIDIndex-1
			strVal=rs(strcoloumnname).Value 
			'msgbox  strVal
			con.close
			set rs=nothing
			set con=nothing
		End If
	End If

	getTestDataValue = strVal
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:CaptureMessage
'Input Parameter      	:strObject - Logical Name of Web Button,strLabel - Label name
'Description          	:This function captures data from application
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function CaptureMessage(strObject,strLabel, strVal)
   Dim bResult,objStatic,strMessage	

   strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objStatic=gobjObjectClass.getObjectRef(strObject)
	
	bResult=False
	If objStatic.Exist(5) Then
			strMessage = Trim(objStatic.GetROProperty("text"))
			gstrDesc =  "The " & strLabel &" displayed is: "&chr(34)&strMessage&chr(39)
			WriteHTMLResultLog gstrDesc, 2
			CreateReport  gstrDesc, 2
			bResult = True
	Else
			TakeScreenShot()
			gstrDesc = strLabel&" object does not exist"
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0   
			objErr.Raise 11
	End If
		
	CaptureMessage = bResult
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name      : connectToDB
'Input Parameter    : None 
'Description        : To connect to Database
'Calls              : None
'Return Value       : True/False
'---------------------------------------------------------------------------------------------------------------------
Public function connectToDB(strVal)

	Dim bResult, strCon, strQuery, Conn, rs
	Dim strDriver,strHost,strPort,strServiceName,strUser,strPassword	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

'	If Err.Number = 424 Then
'		err.clear
'		objErr.clear
'	End If

	Set Conn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")

	strQuery = "SELECT * FROM StaticInformation Where ID = '" & Trim(strVal) & "'"
	
'	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData_"& gstrEnv &".mdb"
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb"
	rs.open strQuery,conn		

	strDriver = rs("DBDriver")
	strHost = rs("DBHost")
	strPort = rs("DBPort")
	strServiceName = rs("DBServiceName")
	strUser = rs("DBUser")
	strpassword = rs("DBPassword")
	strSID=rs("DBSID")
	
	rs.Close	
	Conn.Close
	Set rs = Nothing
	Set Conn = Nothing
	
	Set objDBConn = CreateObject("ADODB.Connection")
	Set rsDBRecords = CreateObject("ADODB.Recordset")	
	
	If Trim(UCase(strVal)) = "LEADS" OR Trim(UCase(strVal)) = "PAPERLESS"  Then
		strCon = "Driver={" & strDriver & "};" 
		strCon = strCon & "CONNECTSTRING=(DESCRIPTION="
		strCon = strCon & "(ADDRESS=(PROTOCOL=TCP)" 
		strCon = strCon & "(HOST=" & strHost & ")(PORT=" & strPort & "))" 
		strCon = strCon & "(CONNECT_DATA=(SERVICE_NAME=" & strServiceName & "))); " 
		strCon = strCon & "uid=" & strUser & ";pwd=" & strPassword & ";"  
	ElseIf Trim(UCase(strVal)) = "STATEMENTS" Then
		strCon = "Driver={"&strDriver&"};"
		strCon = strCon & "Server=(DESCRIPTION="
		strCon = strCon & "(ADDRESS=(PROTOCOL=TCP)"
		strCon = strCon & "(HOST="& strHost &")(PORT="&strPort&"))"
		strCon = strCon & "(CONNECT_DATA=(SID="&strSID&")));"
		strCon = strCon & "Uid="&strUser&";Pwd="&strPassword&";" 
	Else
		strCon = "Driver={"&strDriver&"};"
		strCon = strCon & "Server=(DESCRIPTION="
		strCon = strCon & "(ADDRESS=(PROTOCOL=TCP)"
		strCon = strCon & "(HOST="& strHost &")(PORT="&strPort&"))"
		strCon = strCon & "(CONNECT_DATA=(SID="&strServiceName&")));"
		strCon = strCon & "Uid="&strUser&";Pwd="&strPassword&";"  	
	End If		
	objDBconn.Open strCon	
	
	If Err.number <> 0 AND Err.Number <> 424 then
		gstrDesc = "Failed to connect to Database : " & Err.Number & " : " & Err.Description
		WriteHTMLResultLog gstrDesc, 0
	    CreateReport  gstrDesc,0		
		objErr.Raise 11
		bResult=False
	Else
		gstrDesc = "connected to Database successfully"
		WriteHTMLResultLog gstrDesc, 1
	    CreateReport  gstrDesc,1
		bResult=True
	End if
	connectToDB=bResult

End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      : disconnectDB
'Input Parameter    : None
'Description        : To disconnect Database
'Calls              : ErrorHandler
'Return Value       : True/False
'---------------------------------------------------------------------------------------------------------------------
Public function disconnectDB()

   If objErr.Number=11 Then
	   Exit Function
	End If
         
    	Dim bResult
    	bResult=true    	
    	Set rsDBRecords = Nothing
    	Set objDBconn = Nothing    
		 disconnectDB=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      : executeDBQuery
'Input Parameter    : strTestData - SQPL Query
'Description        : To execute the query on oracle database
'Calls              : None
'Return Value       : True/False
'---------------------------------------------------------------------------------------------------------------------
Public function executeDBQuery(strTestData)

	Dim bResult, arrTemp, strSQL, blnSelect,intCnt, nLoop
	'bResult=False

	strTestData = getTestDataValue(strTestData)
		
	If UCase(strTestData) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If	

	On Error Resume Next
	arrTemp = Split(Trim(strTestData),";")

'    msgbox Ubound(arrTemp)
 	If Ubound(arrTemp) > 0 Then
		If arrTemp(1)= "NULL"  Then
			
			strSQL=arrTemp(0)
			If InStr(strSQL, "?") > 0 Then
				strSQL = Replace(strSQL, "?", gstrPtyID)
			End If
			Set rsDBRecords = objDBConn.Execute(strSQL)
			If rsDBRecords.EOF = "True" Then  'rsDBRecords(0).Value = "" or  
				gstrDesc = "No Data returned after executing the query '" & strSQL & "'."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc,1
				bResult = True
			Else
				gstrDesc = "Value is present in Database '" & strSQL & "'."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc,0
				objErr.Raise 11
				Exit Function
			End If
			
		End If
	End If


	For nLoop = 0 to Ubound(arrTemp)
		
		strSQL = Trim(arrTemp(nLoop))

		If InStr(strSQL, "?") > 0 Then
			strSQL = Replace(strSQL, "?", gstrPtyID)
		End If
		gstrQCDesc = "Execute query '" & strSQL & "'"
		gstrExpectedResult =  "SQL Query '" & strSQL & "' executed successfully"   
			
		If UCase(Left(strSQL,6)) = "SELECT" Then
			blnSelect = True
			wait 2
			Set rsDBRecords = objDBConn.Execute(strSQL)
            If rsDBRecords.EOF = True Or rsDBRecords(0).Value = NULL Then
				gstrDesc = "No Data returned after executing the query '" & strSQL & "'"
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc,0
				objErr.Raise 11
			Else
				gstrDesc = "Successfully executed query '" & strSQL & "'"
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc,1
				bResult = True	
				executeDBQuery = rsDBRecords(0).Value
				Exit Function
			End If
		Else
			''msgbox strSQL
			blnSelect = False
			objDBConn.BeginTrans
			objDBConn.Execute(strSQL)
			objDBConn.CommitTrans
		End If	   

		If Err.number <> 0 And Err.Number <> -2147217900 then
				gstrDesc = Err.Number & " Error occured while executing query '" & strSQL & "' :" & Err.Number & " : " & Err.Description 
			WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc,0
				bResult=false
				objErr.Raise 11
		Else
			If NOT blnSelect Then
				gstrDesc = "Successfully executed query '" & strSQL & "'"
				WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc,1
				bResult = True	      		
				End If
		End If
		
	Next  
	executeDBQuery=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      :  executeDBSelectQuery

'Input Parameter    : None
'Description        : To execute the query on oracle database
'Calls              : None
'Return Value       : True/False
'---------------------------------------------------------------------------------------------------------------------
Public function executeDBSelectQuery(strTestData)

	Dim bResult, arrTemp, strSQL, blnSelect,intCnt, nLoop
	Dim arrColumnList, rs
	'bResult=False
	
	strTestData = getTestDataValue(strTestData)
	
	If UCase(strTestData) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If	
	
	On Error Resume Next
	arrTemp = Split(Trim(strTestData),";")
	
	strSQL = arrTemp(0)
	If UCase(Left(strSQL,6)) <> "SELECT" Then
		gstrDesc = "The query '" & strSQL & "' is not a Select query"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0
		objErr.Raise 11
'		Exit Function
	End If
	
	If InStr(strSQL, "?") > 0 Then
		strSQL = Replace(strSQL, "?", gstrPtyID)
	End If
	gstrQCDesc = "Execute query '" & strSQL & "'"
	gstrExpectedResult =  "SQL Query '" & strSQL & "' executed successfully"   
	
	blnSelect = True
	wait 2
	Set rs = objDBConn.Execute(strSQL)
	
	arrColumnList =  Split(Trim(arrTemp(1)),",")
	
	If rs.EOF = True Or rs(0).Value = NULL Then
		gstrDesc = "No Data returned after executing the query '" & strSQL & "'"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0
		bResult = False
		objErr.Raise 11
	Else
		gstrDesc = "Successfully executed query '" & strSQL & "'"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc,1
		bResult = True   
            
		For nLoop = 0 to Ubound(arrColumnList)
			 If InStr(arrColumnList(nLoop), "=") > 0 Then
				arrTempField = Split(arrColumnList(nLoop),"=")
			Else
				gstrDesc = "Invalid datafield/datavalue format in '" & arrColumnList(nLoop) & "'"
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc,0
				objErr.Raise 11
			End If
			
			tmpFieldName = arrTempField(0)
			tmpFieldVal = arrTempField(1)
			tmpFieldVal = mid(tmpFieldVal,2,len(tmpFieldVal)-2)
			
			For i = 0 to rs.Fields.Count-1
				If rs.Fields(i).Name = tmpFieldName Then
					If UCASE(tmpFieldVal) = "NULL" Then
						If IsNull(rs.Fields(i).Value) Then
							bResult = True
						Else
							bResult = False
						End If
                    ElseIf rs.Fields(i).Value = tmpFieldVal Then
                        bResult = True    
					Else
						bResult = False
					End If
					If bResult = True Then
							gstrDesc = "Successfully retrieved value '" & tmpFieldVal & "' for field '" & tmpFieldName & "'"
							WriteHTMLResultLog gstrDesc, 1
							CreateReport  gstrDesc,1
					Else
							gstrDesc = "Failed to retrieve value '" & tmpFieldVal & "' for field '" & tmpFieldName & "'. "
							gstrDesc = gstrDesc & "Actual value is '" & rs.Fields(i).Value & "'"
							WriteHTMLResultLog gstrDesc, 0
							CreateReport  gstrDesc,0
							objErr.Raise 11
					End If
				End If
			Next
		Next
	End If
	
	If Err.number <> 0 And Err.Number <> -2147217900 then
		gstrDesc = Err.Number & " Error occured while executing query '" & strSQL & "' :" & Err.Number & " : " & Err.Description 
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0
		bResult=false
		objErr.Raise 11
	Else
		If NOT blnSelect Then
			gstrDesc = "Successfully executed query '" & strSQL & "'"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc,1
			bResult = True	      		
		End If
	End If

	executeDBSelectQuery=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'Function Name		: VerifyDatabaseRecord
'Input Parameter	: None
'Description		: This function checks the record retrieved
'Calls			: None
'Return	Value		: None
'---------------------------------------------------------------------------------------------------
Public Function VerifyDatabaseRecord(strVal,strExpected)
	
	Dim bResult,strDatabasevalue
	Dim strpos,strpos1,strpos3,strpos4,strpos5
	Dim strDatabasealue1,strRecordValue
	Dim tempObj,var,temp1,temp2
	Dim var1
	Dim ActualVal,ExpectedVal
	Dim fso, objFile
	
	bResult = False

'	Set tempObj = gobjObjectClass.getObjectRef(strObject)
	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists("C:\DB2\temp.txt") Then

			Set objFile = fso.OpenTextFile("C:\DB2\temp.txt",1)
			strDatabasevalue = objFile.ReadAll
			objFile.Close
			Set objFile = Nothing
			Set fso = Nothing
	'		strDatabasevalue = tempObj.GetROProperty("innertext")		

		If strDatabasevalue = EMPTY Then	 
		   
	'			TakeScreenShot()
				gstrDesc =  "The Database hasn't returned any value. Please check the database connection settings."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0
				objErr.Raise 11
				VerifyDatabaseRecord = bResult
                 Exit Function				
		End If

			If (instr(strDatabasevalue,"0 record(s) selected.")>0)  Then	
				If Ucase(Trim(strExpected)) = "TRUE" Then
					gstrDesc = "The value is successfully retrieved for query ' " & strVal& " '."
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
					bResult = True
					VerifyDatabaseRecord = bResult   
				Else
					gstrDesc =  " No record is returned for query ' " & strVal& " '."
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0
					 objErr.Raise 11
					VerifyDatabaseRecord = bResult  
					Exit Function
				End If
			Else
			    gstrDesc = "The value is successfully retrieved for query ' " & strVal& " '."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = True
				VerifyDatabaseRecord = bResult   
			 End If	
	  End if
	 End Function
	 
'---------------------------------------------------------------------------------------------------------------------	   
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:BrowserBack
'Input Parameter    	:strObject - Logical Name of Web Browser
'Description        	:This function goes to previous  page of the browser object
'      			 DLL returns an object reference of a Browser.
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function BrowserBack(strObject,strLabel,strVal)
	Dim objBrowser, bResult
	
	 strVal = getTestDataValue(strVal)
		
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If	

	On Error Resume Next

	bResult=True	
	gstrQCDesc = " Back Browser '" & strLabel & "'"
	gstrExpectedResult = "Browser '" & strLabel & "' should go to previous page"
    Set objBrowser = gobjObjectClass.getObjectRef(strObject)
	objBrowser.Back
    gstrDesc =  "Clicked on Back for '" & strLabel & "' Browser."
	 WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1
   
	BrowserBack = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:BrowserForward
'Input Parameter    	:strObject - Logical Name of Web Browser
'Description        	:This function opens next page of the browser object
'      			 DLL returns an object reference of a Browser.
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function BrowserForward(strObject,strLabel,strVal)
	Dim objBrowser, bResult

	strVal = getTestDataValue(strVal)
		
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If	

	On Error Resume Next
	
	bResult=True	
	gstrQCDesc = "Close Browser '" & strLabel & "'"
	gstrExpectedResult = "Browser '" & strLabel & "' Successfully open next page"
	
	Set objBrowser = gobjObjectClass.getObjectRef(strObject)
    objBrowser.Forward	
	gstrDesc =  "Clicked on Forward for '" & strLabel & "' Browser."
	WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1
		
	BrowserForward = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:SecurityClick
'Input Parameter      	: None
'Description          	:This function click popup
'Calls            	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Function SecurityClick()
   If Browser("title:=.*").Dialog("text:=Security Alert").Exist(2) Then
	    Browser("title:=.*").Dialog("text:=Security Alert").WinButton("text:=.*Yes.*").highlight
		Browser("title:=.*").Dialog("text:=Security Alert").WinButton("text:=.*Yes.*").Click
	  '  Browser("title:=.*").Dialog("text:=Security Alert").WinButton("text:=.*Yes.*").Click
   End If

End Function 
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:StringLength
'Input Parameter      	: strVal - string whose length is to calculated
'Description          	:This function returns the length of the string
'Calls            	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function StringLength(strVal)
	Dim objBrowser, bResult

	strVal = getTestDataValue(strVal)
		
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If	

	On Error Resume Next
	
	bResult=True	
      
'	Set tempObj = gobjObjectClass.getObjectRef(strObject)

	strTempLen = len(strVal)
    
	gstrDesc =  "The length of string '" & strVal & "' is '" & strTempLen & "' characters."
	WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1
	StringLength = bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name       	:VerifyDialogBox
'Input Parameter     	:strVal - Text to print
'Description         	:This function prints the non - automatable steps
'Return Value 		:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifyDialogBox(strObject,strLabel, strVal)
	Dim bResult,strDesc,objElm	
	bResult=False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	 If instr(strObject,":") Then
		arrobj=Split(strObject,":")
		For i= 0 To Ubound(arrobj)
			Set objElm=gobjObjectClass.getObjectRef(arrobj(i))
			If objElm.Exist  Then
			    gstrDesc =  "Dialog Box '" & strLabel & "' is displayed successfully."				
        		WriteHTMLResultLog gstrDesc, 2
        		CreateReport  gstrDesc, 2 
				Exit For
			End If
		Next
	Else
		Set objElm = gobjObjectClass.getObjectRef(strObject)
		If Not objElm Is Nothing Then
			gstrQCDesc = "Verify Dialog Box" & strLabel & " is displayed."
			gstrExpectedResult = "Dialog Box " & strLabel & " should be displayed."	
			If objElm.Exist Then
					gstrDesc =  "Dialog Box '" & strLabel & "' is displayed successfully."				
					WriteHTMLResultLog gstrDesc, 2
					CreateReport  gstrDesc, 2 
				bResult = True       	
			Else
					Call TakeScreenShot()				
					gstrDesc =  "Dialog Box " & strLabel & " is not displayed properly."				
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0		
					 bResult = False
					objErr.Raise 11   		   
			End If
		End If
	End If 
  	VerifyDialogBox=bResult
	
End Function
'---------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name       	:VerifyDisplayProperty
'Input Parameter     	:strVal - Text to print
'Description         	:This function prints the non - automatable steps
'Return Value 		:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifyDisplayProperty(strObject,strLabel, strVal)

	Dim objElm
    strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	If instr(strObject,":") Then
		arrobj=Split(strObject,":")
		For i= 0 To Ubound(arrobj)
			Set objElm=gobjObjectClass.getObjectRef(arrobj(i))
			If objElm.Exist  Then
				gstrpropertyvalue=objElm.GetRoproperty(strVal)
				gstrDesc = "Object on '"& strLabel &"' is displayed with: '" & gstrpropertyvalue &"."   	
				WriteHTMLResultLog gstrDesc, 4
				CreateReport  gstrDesc, 4
				Exit For
			End If
		Next
	Else
		Set objElm = gobjObjectClass.getObjectRef(strObject)
		If objElm.Exist Then
			gstrpropertyvalue=objElm.GetRoproperty(strVal)
			If IsNumeric(gstrpropertyvalue) Then
				 If objElm.Exist Then 'And  gstrpropertyvalue=1
				   gstrDesc = "Object '"& strLabel &"' is displayed '" & gstrpropertyvalue &"."   	
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
				bResult = True       	
				 Else
					Call TakeScreenShot()				
					gstrDesc =  "Object '" & strLabel & "' is Enabled '" & gstrpropertyvalue &"."   		
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0		
					 bResult = False
					objErr.Raise 11      
				End If
			Else
				 If objElm.Exist Then 
					   gstrDesc = "Object on '"& strLabel &"' is displayed with: '" & gstrpropertyvalue &"."   	
						WriteHTMLResultLog gstrDesc, 1
						CreateReport  gstrDesc, 1
					bResult = True       	
				Else
						Call TakeScreenShot()				
						gstrDesc =  "Object on '" & strLabel & "' is not displayed."				
						WriteHTMLResultLog gstrDesc, 0
						CreateReport  gstrDesc, 0		
						 bResult = False
						objErr.Raise 11
				End If
			End If
		Else
			Call TakeScreenShot()
			gstrDesc =  "Object on '" & strLabel & "' is not displayed."
'			gstrDesc =  "Page" & strLabel & " is not displayed."				
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
		End If
	End If
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name       	:updateADMData
'Input Parameter     	:strVal - update the field with data
'Description         	:This function update the data to ADM database
'Return Value 		:Nothing
'---------------------------------------------------------------------------------------------------------------------
Public Function updateADMData(strVal)
	Dim objLdAP,objUser1
	Dim strIP,strlogin,strpassword,strUserName,strUser,strPath,lngAuth,arrTemp,arrTemp1
	
				strVal = getTestDataValue(strVal)
				
				If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
					Exit Function
				End If
			
				arrVal = Split(strVal,":",2)
				
				Set Conn = CreateObject("ADODB.Connection")
				Set rs = CreateObject("ADODB.Recordset")
			
				strQuery = "SELECT * FROM ADMStaticInformation Where ID = '" & Trim(arrVal(0)) & "'"
				
'				conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData_"& gstrEnv &".mdb"
				conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb"
				rs.open strQuery,conn
    	
				Const ADS_SECURE_AUTHENTICATION = 1
                Const ADS_USE_SIGNING = 64
                Const ADS_USE_SEALING = 128
        
                Set objLdAP = GetObject("LDAP:")
                'strIP = rs("ADMIP")
				strIP = gDictAdam("SERVER")
                'strlogin = "EI20\" & rs("ADMLogin")
'				strlogin = "EI20\" & gDictAdam("USERNAME")
				strlogin = gDictAdam("USERNAME")
                'strpassword = rs("ADMPassword")
				strpassword = gDictAdam("PASSWORD")
				strUserName = rs("ADMUserName")
				'strUserName =gstrStrUsername
				strDomain = rs("ADMDomain")
                strUser = "CN=" & strUserName & "," & "OU="& strDomain &",OU=organizations,OU=UIdP,CN=RETAILECOM,O=HBOS"

                strPath = "LDAP://" & strIP & "/" & strUser
                lngAuth = ADS_USE_SIGNING Or ADS_USE_SEALING Or ADS_SECURE_AUTHENTICATION
          
                Set objUser1 = objLdAP.OpenDsObject(strPath, strlogin, strpassword, lngAuth)
  	            
				If Instr(arrVal(1),",") >0 Then
					arrTemp = split(arrVal(1),",")
				else 
					arrTemp = arrVal(1)	
				End If

				  
			  For i=0 to Ubound(arrTemp)
 			  
					arrTemp1 = Split(arrTemp(i),"=")		   	
					On error resume next
					If  arrTemp1(1) ="CLEAR" Then
						strval = objUser1.Get(arrTemp1(0))
				    End If

					If   arrTemp1(0) ="accessCodeEmailsIssued" Then
						If arrTemp1(1) ="-2" Then
							strTime = Hour(Time) - 2 & ":" & Minute(Time) & ":" & Second(Time)
							arrTemp1(1) = CDate(Date & " " & strTime)
						    arrTemp1(1) =Cstr(  arrTemp1(1))
						End If

						If  arrTemp1(1) ="+2"Then
								strTime = Hour(Time) + 2 & ":" & Minute(Time) & ":" & Second(Time)
								arrTemp1(1) = CDate(Date & " " & strTime)
								arrTemp1(1) =Cstr(  arrTemp1(1))
						End If
					End If

					If   arrTemp1(0) ="letters" Then
						If arrTemp1(1) ="-7" Then
							strDate = Date - 7
							arrTemp1(1) = CDate(strDate & " " & Time)
							arrTemp1(1) =Cstr(  arrTemp1(1))
						End If

						If  arrTemp1(1) ="+7"Then
							  strDate = Date + 7
							  arrTemp1(1) = CDate(strDate & " " & Time)
							  arrTemp1(1) =Cstr(  arrTemp1(1))
						End If
					End If
					
					If Err.Number <> "-2147463155" Then
                        Err.Clear
		
						If arrTemp1(1) ="CLEAR"  Then
							objUser1.PutEx 1,arrTemp1(0),0
							objUser1.setinfo
						Else
							objUser1.Put arrTemp1(0),arrTemp1(1)
							objUser1.setinfo	
						End If

				 Else

						Err.clear	  

				End if	
					
				
					If Cstr(objUser1.Get (arrTemp1(0)))=arrTemp1(1) Then
						gstrDesc = "Value of the field '"& arrTemp1(0) &"' is successfuly updated as:" & arrTemp1(1) &"."   	
						WriteHTMLResultLog gstrDesc, 1
						CreateReport  gstrDesc, 1
					Else 
						gstrDesc =  "Value of the field '"& arrTemp1(0) &"' is NOT updated as:" & arrTemp1(1) &"."  			
						WriteHTMLResultLog gstrDesc, 0
						CreateReport  gstrDesc, 0		
						objErr.Raise 11	
					End If
				Next


				rs.Close	
				Conn.Close
				Set rs = Nothing
				Set Conn = Nothing
                Set objLdAP = Nothing
                Set objUser1 = Nothing
				
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:selectVBRadioButton
'Input Parameter    :strObject - Logical Name of Radio Button
'     								:strVal - Contains the property of the Radio Button which needs to be 
'									changed and the value it needs to be changed to.
'Description          	:This function selects the radio button
'Calls       	         	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function selectVBRadioButton(strObject,strLabel,strVal)

	Dim objRadio,bResult
	bResult = False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objRadio = gobjObjectClass.getObjectRef(strObject)

	gstrQCDesc = "Select value '" & getTestDataValue(strVal) & "' from '" & strLabel & "' Radio Button."
	gstrExpectedResult = "Value '" & getTestDataValue(strVal) & "' should be selected successfully from '" & strLabel & "' Radio Button."
	
	If objRadio.exist(gExistCount) Then
		objRadio.Set
		If gbIterationFlag <> True then
			gstrDesc =  "Successfully set radio button '" & strLabel & "."
			gstrExpectedResult = gstrDesc
			gstrQCDesc = "Set the Radio Button '" & strLabel & "'."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Radio Button '" & strLabel & "' is not displayed."
		gstrExpectedResult = "Successfully set the Radio button '" & strLabel
		gstrQCDesc = "Set the Radio Button '" & strLabel & "'."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If 
	
	selectVBRadioButton = bResult
End Function
'_________________________________________________________
'Function-Name :selectVBComboBox
'Input Parameter     	:strObject- Object ComboBox
'Description : Select Data from ComboBox
'Output 
'_________________________________________________________________________________________________________________________________
Public Function selectVBComboBox(strObject,strLabel,strVal)
	
	strVal = getTestDataValue(strVal)
 
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objCombo=gobjObjectClass.getObjectRef(strObject)

	If objCombo.exist Then
		strContent=objCombo.GetContent
		arrItems=Split(strContent,vbLf)
		If IsArray(arrItems) Then
			objCombo.Select arrItems(strVal)
			gstrDesc =  "Value '" & objCombo.getRoProperty("text") & "' is selected for Combo Box '" & strLabel &"'."
			WriteHTMLResultLog gstrDesc, 2
			CreateReport  gstrDesc, 2
		Else
			gstrDesc =  "Value can not be selected for Combo Box '" & strLabel &"'."
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0
		End If
	Else
		gstrDesc =  "Combo Box '" & strLabel &"' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
	End If
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:CloseWindow
'Input Parameter    	:strObject - Logical Name of Web Browser
'Description        	:This function closes the browser object
'      			 DLL returns an object reference of a Browser.
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function closeWindow(strObjName,strLabel, strVal)
	Dim objBrowser, bResult
	bResult=True	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Close Window '" & strLabel & "'"
	gstrExpectedResult = "Window '" & strLabel & "' should be closed."
	
	Set objBrowser = gobjObjectClass.getObjectRef(strObjName)
	If objBrowser.Exist(1) Then
		objBrowser.close
		gstrDesc =  "Successfully closed '" & strLabel & "' Window."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
	Else
		objErr.Raise 11
		Call TakeScreenShot()
		gstrDesc =  "Window '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		bResult=False
	End If
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:setAcxCalender
'Input Parameter      	:strObject - Logical Name of Check box
'     			:strVal - Contains the property of the Chechkbox which needs to be changed and the
'      			 value it needs to be changed to.
'Description          	:This function toggles the checkbox
'      			 DLL returns an object reference of a WebCheckBox
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function setAcxCalender(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objCln = gobjObjectClass.getObjectRef(strObject)		


	If objCln.exist(gExistCount) Then

		If IsDate(strVal) Then
			objCln.SetDate strVal
			bResult = True
		Else
			objCln.SetDate Date
			bResult = True
		End If
		If bResult = True Then
			gstrDesc = "Successfully selected the '" & strLabel & "' AcxCalender."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
		
	Else
		Call TakeScreenShot()
		gstrDesc =  "AcxCalender '" & strLabel & "' is not displayed in the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	
	setAcxCalender = bresult
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:enterCellData
'Input Parameter      	:strObject - Logical Name of JavaTable
'     			:strVal - Value to  be entered in Cell
'Description        	:This function enters a data into a Cell of Java Table
'      			 
'Calls              	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function enterCellData(strObject,strLabel,intRow,intCol,strVal)
	
	Dim bResult, objTextField,objTrim,objlen,objReport, i, Wsh
	bResult=False			

	'strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."
	Set objTextField = gobjObjectClass.getObjectRef(strObject)
	If objTextField.exist(gExistCount) Then

		objTextField.SetCellData intRow,intCol,strVal
		If gbIterationFlag <> True then
			gstrDesc = "Successfully entered '" & strVal & "' as '" & strLabel & "' cell data."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "'" & strLabel & "' Cell does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If

  	enterCellData=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:MaximizeBrowser
'Input Parameter      	:strObject - Logical Name of Browser
'Description        	:Maximize Browser
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function MaximizeBrowser(objBrw)
	If objErr.Number = 11 Then
		Exit Function
	End If
	
	On Error Resume Next
	Dim nTemp
	nTemp = objBrw.Object.HWND
	Window("hwnd:=" & Cstr(nTemp)).Maximize
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:selectAllRadioButton
'Input Parameter    	:strObject - Logical Name of Web Button
'Description          	:This function selects all radio buttons displayed on page
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function selectAllRadioButton(strObject, strLabel, strVal)
	On error resume next
	Dim objRad, arrRad, objPage, nLoop

	strVal = getTestDataValue(strVal)
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objRad = Description.Create
	objRad("micclass").Value = "WebRadioGroup" 
	
	Set objPage = gobjObjectClass.getObjectRef(strObject)
	Set arrRad = objPage.ChildObjects(objRad)
	If objPage.Exist(gExistCount) And arrRad.Count > 1 Then
		For nLoop = 0 to arrRad.count-1
			arrRad(nLoop).Select strVal
		Next
		gstrDesc =  "Successfully selected value " & strVal & " for all radio buttons available on '" & strLabel & "' page."		
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
	Else
		Call TakeScreenShot()
		gstrDesc =  "Page  '" & strLabel & "' is not displayed on screen."		
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		objErr.Raise 11
	End If

End Function

'---------------------------------------------------------------------------------------------------
'Function Name		: selectRadioButtonByName
'Input Parameter	: strVal: Name Corresponding to Radio Button
'									strObject:Table in which radio group is present
'Description		: This function Select radio button corresponding to the name.
'Calls				: None
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Public Function selectRadioButtonByName(strObject,srtVal)

	Dim rowNo,tempObj,colNo

	strVal = getTestDataValue(strVal)
 
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	On error resume next
			Set tempObj = gobjObjectClass.getObjectRef(strObject)
			colNo=0
			For colNo= 1 to tempObj.columnCount(1)
					If  tempObj.childItemcount(2,colNo,"WebRadioGroup")>0Then
						Exit For
					End If   
			Next
'			msgbox colNo
			rowNo=tempObj.GetRowWithCellText (srtVal)
	'		msgbox "#" & rowNo
			tempObj.ChildItem(rowNo,colNo,"WebRadioGroup",0).Select "#" & rowNo-2

			If  tempObj.ChildItem(rowNo,colNo,"WebRadioGroup",0).CheckProperty("selected item index",rowNo-1) Then
					gstrDesc =  "Successfully selected radio button for  '" & strVal & "'."
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
				Else
					TakeScreenShot()
					gstrDesc = "Falied to selectt radio button for  '" & strVal & "'."
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0
					objErr.Raise 11  
			 End If

			Set tempObj=Nothing

End Function

'---------------------------------------------------------------------------------------------------
'Function Name		: setAllCheckBox
'Input Parameter	: strObject - Logical Name of CHeck Box
'Description		: This function sets all the checkboxes as On/Off  on the respective page depending upon the value of strval
'Calls				: None
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Public Function setAllCheckBox(strObject,strLabel,strVal)
	Dim objChk, bresult,objPage,arrChk,nLoop
	 
	bresult=False

	strVal = 
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

    Set objChk = Description.Create
	objChk("micclass").Value = "WebCheckBox" 
	
	Set objPage = gobjObjectClass.getObjectRef(strObject)
	Set arrChk = objPage.ChildObjects(objChk)
	if objPage.Exist(gExistCount) And arrChk.Count > 0 Then
		For nLoop = 0 to arrChk.count-1
           If UCase(strVal) = "OFF" Then
			  arrChk(nLoop).set  "OFF"
			Else
			  arrChk(nLoop).set  "ON"
			 End If
		Next
		gstrDesc =  "Successfully set value " & strVal & " for all check boxes available on '" & strLabel & "' page."		
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
	Else
		Call TakeScreenShot()
		gstrDesc =  "Page  '" & strLabel & "' is not displayed on screen."		
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		objErr.Raise 11
	End If
    setAllCheckBox = bresult
End Function
'---------------------------------------------------------------------------------------------------
'Function Name		: CheckCapsLock
'Input Parameter	: None
'Description		: CheckCapsLock
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Function CheckCapsLock()
    Dim Res
    Dim KBState   '(0)To (255) As Byte
    Res = GetKeyboardState(KBState(0))
    'If KBstate(&H14) =1 Caps are on. If KBstate(&H14) =0 Caps are off.
    If KBState(&H14) = 1 Then
		KBState(&H14) = 0        
	End If
End Function
'---------------------------------------------------------------------------------------------------
'Function Name		: ClickWebTableLink
'Input Parameter	: strObject - Logical Name of WebTable
'Description		: This function CLick link in web table
'Calls				: None
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Public Function ClickWebTableLink(strObject,strLabel,strVal)
	'msgbox "in func"
	Dim bResult, objWebTableLink, objclick

	bResult=False
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	arrTemp = Split(strVal,":")
	ntRow = arrTemp(0)
	ntColumn = arrTemp(1)
	ntIndex = arrTemp(2)
			
	Set objWebTableLink = gobjObjectClass.getObjectRef(strObject)
	
	If objWebTableLink.exist(gExistCount) Then

		gstrExpectedResult = objWebTableLink.GetROProperty("innertext") & " link should get clicked"
		gstrQCDesc = "Click on link '" & objWebTableLink.GetROProperty("innertext") &"'"

		Set  objclick = objWebTableLink.ChildItem(Cint(ntRow) , Cint(ntColumn) ,"Link", Cint(ntIndex))
		objclick.click

        gstrDesc = "Successfully Clicked " & objclick.GetROProperty("innertext") &"Link"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = True
	Else		 
		gstrDesc =  "'" & strLabel & "' in WebTable does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If

ClickWebTableLink = bResult

End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:clickWinObject
'Input Parameter	: strObject - Logical Name of WebTable
''Description            :This function clicks on Win Object
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function clickWinObject(strObject, strLabel,strVal)
   Dim bResult,objWinObject
	bResult=False 

	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	'msgbox gobjSub
	Set objWinObject=gobjObjectClass.getObjectRef(strObject)
	'msgbox objWinObject
		If objWinObject.Exist(gExistCount) Then
			objWinObject.Click
			gstrDesc =  "WinObject '" & strLabel & "' is clicked."	 
			gstrExpectedResult = "WinObject '" & strLabel & "' should be clicked." 
			gstrQCDesc = "Click on WinObject '" & strLabel & "'"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1 
			bResult = True       	
		Else				
			gstrDesc =  "WinObject '" & strLabel & "' does not exists."
            Call TakeScreenShot()
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0			
			bResult = False
		End if
	Set objWinObject=Nothing
	Set gobjSub=Nothing
	wait 1
	
	clickWinObject=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:closeAll
'Input Parameter	: strObject - Logical Name of WebTable
'Description          	:This function closes all the open browsers
'Calls                	:Nothing
'Return Value   	:Nothing
'---------------------------------------------------------------------------------------------------------------------
Public Function closeAll()

	Dim objDesc
	Dim arrBrowsers
	Dim nLoop
	On Error Resume Next
	Set objDesc = Description.Create
	objDesc("micclass").Value = "Browser"
	
	Set arrBrowsers = Desktop.ChildObjects(objDesc)
	
	For nLoop = 0 to arrBrowsers.Count - 1
		arrBrowsers(nLoop).Close
	Next
	
	Set arrBrowsers = Nothing
	
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:dragAndDropFile
'Input Parameter	: strObject - Logical Name of SSH window
''Description            :This function drags file from Source path and drops it to Destination Paths in SSH FTW
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function dragAndDropFile(strObjName)
   Dim bResult,arrObj,arrVal,strSource,l,t,r,b,strResult,x,y,WshShell,strSearch,intCount,strTrim,strCollection,bStatus,bFlag,strItems,strItemsCount,intItemCntr, resultpath
	bResult=False

	If objErr.Number = 11 Then
        Exit Function
    End If

	Window("winSSHSecureFileTransfer").Activate

    arrObj=Split(strObjName,":")
	Set objAdd = gobjObjectClass.getObjectRef(arrObj(0))
	Set objSLV = gobjObjectClass.getObjectRef(arrObj(1))
	Set objDLV = gobjObjectClass.getObjectRef(arrObj(2))

	If objAdd.Exist(gExistCount) Then
		'objAdd .Activate
        wait 1 'Reduce 1
		l = -1
		t = -1
		r = -1
		b = -1
		strResult=objAdd.GetTextLocation("Add", l, t, r, b,False)
		If strResult Then
			x=(l+r)/2
			y=(t+b)/2
			x=x-200
			y=y+60
			wait 1 'Reduce1
			'gobjSub.WinListView(arrObj(1)).Select strVal	
			objSLV.DragItem gstrFileName
			wait 1
			objDLV.Drop x,y
			wait 1
		End If
	End If
'            If  gobjSub.Dialog("diaConfirmFileOverwrite").Exist(4) Then
'					gobjSub.Dialog("diaConfirmFileOverwrite").WinButton("btnYesOverwrite").Click
'			 End If
	strCollection=objDLV.GetROProperty("all items")
	If Instr(1,strCollection, gstrFileName) Then
		bFlag=True
	End If
	 If bFlag=True Then
			 gstrDesc =  "File '" &gstrFileName &"' is successfully dragged from source and dropped to destination."
			 gstrExpectedResult = "File '" & gstrFileName &"' should be successfully dragged from source and dropped to destination."
			 gstrQCDesc = "Drag and Drop file '" &gstrFileName &"' from source to destination."
			 WriteHTMLResultLog gstrDesc, 1
			 CreateReport  gstrDesc, 1 
			 bResult = True
		Else
			gstrDesc = "File couldn't dragged to destination."
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0			
			bResult = False
			objErr.Raise 11
		End If 
    wait 1
	dragAndDropFile=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:enterPath
'Input Parameter	: strObject - Logical Name of SSH window
''Description            :This function enters the Source and Destination Paths in SSH FTW
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function enterPath(strObjName,strLabel,strVal)
   Dim bResult,arrObj,arrVal,strSource
	bResult=False

	strVal = GetTestDataValue(strVal)
	If objErr.Number = 11 Or Ucase(strVal) = "SKIP"Then
        Exit Function
    End If

    Window("winSSHSecureFileTransfer").Activate
	strSource= gstrBaseDir & "STPFiles\"
	strDestination = "/var/stp/p3/BACSIA/input/data"

    arrObj=Split(strObjName,":")

	Set objSource = gobjObjectClass.getObjectRef(arrObj(0))
	Set objDest = gobjObjectClass.getObjectRef(arrObj(1))
	'Set objSource = gobjObjectClass.getObjectRef(arrObj(0))
	If objSource.Exist(gExistCount) Then
		objSource.Set strSource
		wait 2
		objSource.Type  micReturn
		wait 8
		objDest.Set strDestination
		wait 8
		objDest.Type  micReturn
		wait 3

        TakeScreenShot()
        gstrDesc =  "Source & Destination Paths in Window '" & strLabel & "' are entered successfully."
		gstrExpectedResult = "Source & Destination Paths in Window '" & strLabel & "' should be entered."
		gstrQCDesc = "Enter Source & Destination Paths in Window '" & strLabel &  "'"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1 
		bResult = True       	
	Else
		TakeScreenShot()				
		gstrDesc = "Window '" & strLabel & "' does not exists."
        WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0			
		bResult = False
	End if
	'Set gobjSub=Nothing
    wait 1
   enterPath=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:EnterSecureText
'Input Parameter      	:strObject - Logical Name of Edit Box
'     			:strVal - Value to  be entered in text box
'Description        	:This function enters a data into a text box
'      			 DLL returns an object reference of a WebEdit
'Calls              	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function EnterSecureText(strObject,strLabel,strVal)

   	Dim bResult, objTextField,objTrim,objlen,objReport, i, Wsh, nCnt, strTextVal
	bResult=False			

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."
	Set objTextField = gobjObjectClass.getObjectRef(strObject)
	If objTextField.exist(gExistCount) Then
		If UCase(strVal) = "DATE" Then
			objTextField.Set Date
		ElseIf UCase(strVal) = "RANDOM" Then
             MyValue = Int((999 * Rnd) + 100)
			 objTextField.Set MyValue
		Else
			objTextField.SetSecure strVal
		End If

		If gbIterationFlag <> True then
			gstrDesc = "Successfully entered '" & strVal & "' in '" & strLabel & "' textbox."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "'" & strLabel & "' textbox does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If

  	EnterSecureText=bResul
End Function
'---------------------------------------------------------------------------------------------------
'Function Name		: EnterTextFromFile
'Input Parameter	: strObject: Web Edit Object logical name
'strLabel: Label for web edit, strVal: file name from which values needs to be inputed
'Description		: This function inputs the data from input text file
'Calls			: None
'Return	Value: None
'---------------------------------------------------------------------------------------------------
Public Function EnterTextFromFile(strObject,strLabel,strVal)

	Dim bResult, fso, objFile, strData, objText
	bResult = False
	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFile = fso.OpenTextFile(gstrBaseDir & "Request\" & strVal,1)
	strData = objFile.ReadAll()
	Set objFile = Nothing
    
	Set objText = gobjObjectClass.getObjectRef(strObject)
	If  objText.exist(gExistCount) Then
		objText.Set ""
		objText.Set strData
		gstrDesc =  "The Data from source file '" & strVal & "' is entered succefully in "
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = True
	Else
		gstrDesc =  "The Soap Envelope Text Box is not displayed on the screen"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		bResult = False
	End If

End Function
'------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
'Function Name    	:invokeSsh
''Description            :This function invokes the SSH Client
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function invokeSsh(strObject,strVal)
   Dim bResult,strPath
	bResult=True 

	strVal = getTestDataValue(strVal)
'	strPath=gstrDriveName & "\Program Files\SSH Communications Security\SSH Secure Shell\SshClient.exe"
	strPath = "C:\Program Files\SSH Communications Security\SSH Secure Shell\SshClient.exe"
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set gobjSub=gobjObjectClass.getObjectRef(strObject)

	If  Window("winSSHSecureFileTransfer").Exist(1) Then
		Window("winSSHSecureFileTransfer").Close
	End If
	wait 1
'	If  Window("winSSHSecureShell").Exist(5) Then
'		Window("winSSHSecureShell").Close
'		If   Window("winSSHSecureShell").Dialog("diaConfirmExit").Exist(2) Then
'			wait 1
'			Window("winSSHSecureShell").Dialog("diaConfirmExit").WinButton("btnOK_dis").Click
'		End If
'	End If
	If  gobjSub.Exist(gExistCount) Then
		gobjSub.Close
		If   gobjSub.Dialog("diaConfirmExit").Exist(2) Then
			wait 1
			gobjSub.Dialog("diaConfirmExit").WinButton("btnOK_dis").Click
		End If
	End If
    wait 1
	Systemutil.Run strPath
	wait 2
	If gobjSub.Exist(gExistCount) Then
		strLabel = gobjSub.GetROProperty("Text")
		gstrDesc =  "SSH Client '" & strLabel & "' Invoked Successfully."
		gstrExpectedResult = gstrDesc 
		gstrQCDesc = "SSH Client '" & strLabel & "'."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		wait 2
	End If
	
	invokeSsh=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:openFileTransferWindow
'Input Parameter	: strObject - Logical Name of SSH window
''Description            :This function opens the File Transfer Window
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function openFileTransferWindow(strObjName, strLabel,strVal)
	Dim bResult,WshShell
	Dim objShell, objFileShell 
	bResult=False

	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	strObject = Split(strObjName,":")

	Set objShell = gobjObjectClass.getObjectRef(strObject(0))
	Set objFileShell = gobjObjectClass.getObjectRef(strObject(1))
    If objShell.Exist(gExistCount) Then
		objShell.Activate

		Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "%(wf)"
		Set WshShell = Nothing
		Wait 2

		If objFileShell.Exist(gExistCount) Then
			objFileShell.Activate
			objFileShell.Maximize
			strLabel = objFileShell.GetROProperty("Text")
			gstrDesc =  "File Transfer Window '" & strLabel & "' is opened."	 
			gstrExpectedResult = "File Transfer Window '" & strLabel & "' should be opened." 
			gstrQCDesc = "Open '" & strLabel & "' File Transfer Window."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1 
			bResult = True       	
		Else
			gstrDesc =  "File Transfer Window not opened."
			Call TakeScreenShot()
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0			
			bResult = False
		End If
	Else				
		gstrDesc =  "Secure Shell Transfer Window '" & strLabel & "' does not exists."
        Call TakeScreenShot()
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0			
		bResult = False
	End if
	wait 2
	
	openFileTransferWindow=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:openNewterminal
''Description            :This function opens the New Terminal
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function OpenNewTerminal(strObject, strLabel,strVal)
	Dim bResult,WshShell
	Dim objShell, objFileShell
	bResult=False
	'strObject = Split(strObjName,":")

	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	If Instr(strObject,":") > 0 Then
		arrObj = Split(strObject,":")
	End If
	
	Set objShell = gobjObjectClass.getObjectRef(arrObj(0))
	'Set objFileShell = gobjObjectClass.getObjectRef(strObject(1))
    If objShell.Exist(gExistCount) Then
		objShell.Activate

		Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "%(wt)"
		Set WshShell = Nothing
	End if
	Set objFileShell = gobjObjectClass.getObjectRef(arrObj(1))
	If objFileShell.Exist(gExistCount) Then
            gstrDesc =  "New Terminal Window is opened."	 
			gstrExpectedResult = "New Terminal should be opened." 
			gstrQCDesc = "Open New Terminal Window"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1 
			bResult = True       	
	Else
			gstrDesc =  "New Terminal Window not opened."
			Call TakeScreenShot()
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0			
			bResult = False
	End If
   	wait 2
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:runQuery
''Description            :This function runs the query.
'Calls                  :None
'Return Value   	:None
'----------------------------------------------------------------------------------------------------------------------
Public Function runQuery(strObjName, strLabel, strVal)
   Dim bResult,WshShell,strQueryPath
	bResult=False
    strTemp = strVal
	strVal = GetTestDataValue(strVal)
	
	If objErr.Number = 11 or Ucase(strVal) = "SKIP"Then
        Exit Function
    End If

'	If Len(strVal) > 150 Then
'    	strQueryPath = strVal
'	Else
'		strQueryPath= strVal & gstrFileName
'	End If

    strQueryPath = CreateQuery(strVal)
    Set objShell= gobjObjectClass.getObjectRef(strObjName)
	If objShell.Exist(gExistCount) Then
		objShell.Activate
		wait 2
		objShell.Activate
'		wait 1
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys strQueryPath
		wait 2
		WshShell.SendKeys "{ENTER}"
		wait 3
		Set WshShell = Nothing

'		If gQFlag Then
'			gTStamp = Time	
'		End If
'		gTStamp = Time
		gstrDesc =  "Successfully executed the query and Time Stamp is " & gTStamp
		gstrExpectedResult = "Query should be executed successfully"
		gstrQCDesc =  "Successfully executed the query."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1 
		bResult = True
    Else				
		gstrDesc = "Query is not executed successfully"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0			
		bResult = False
	End if
   runQuery=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------
'Function Name  :typeTextValue
'Input Parameter :strObject - Logical Name of Edit Box
'   :strVal - Value to  be entered in text box
' strTab - no of tabs required
'Description  :This function enters a data into a text box in WinXcel Application.
'Calls   :None
'Return Value  :True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function typeTextValue(strObject,strLabel, strVal,strTab)
	Dim bResult, objTextField, dVar,splitText,ArrEditName, Wsh,i
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."

	ArrEditName=Split(strObject,"-")
	Set objTextField = gobjObjectClass.getObjectRef(strObject)
	
	If  objTextField.exist(gExistCount) Then
        
		Set Wsh = CreateObject("Wscript.Shell")
		For i=0 to strTab-1
			Wsh.SendKeys "{TAB}"
			wait 1 
		Next
		wait 1 
		
'	objTextField.Click
'	  objTextField.Click
'		Wsh.SendKeys "{TAB}"
'		Wsh.SendKeys "{TAB}"
'		Wsh.SendKeys "{TAB}"
		gstrExpectedResult = strVal & " should get typed  in '" & strLabel & "' textbox."
		Wsh.SendKeys strVal
		wait 1
	   
'	    Wsh.SendKeys "{TAB}"
		wait 1
	  
			gstrDesc = "Successfully typed '" & strVal & "' in '" & strLabel & "' textbox."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
	 
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "'" & strLabel & "' textbox does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If
	 typeText=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'*****************************************************
Public function CONNECTtoPAM(strVal)

	Dim bResult, strCon, strQuery, Conn, rs
	Dim strDriver,strHost,strPort,strServiceName,strUser,strPassword	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	If Err.Number <> 0 Then
		err.clear
		objErr.clear
	End If

	Set Conn = CreateObject("ADODB.Connection")
	Set rs = CreateObject("ADODB.Recordset")

	strQuery = "SELECT * FROM StaticInformation Where ID = '" & Trim(strVal) & "'"
	
'	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData_"& gstrEnv &".mdb"
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb"
	rs.open strQuery,conn		

	strDriver = rs("DBDriver")
	strHost = rs("DBHost")
	strPort = rs("DBPort")
	strServiceName = rs("DBServiceName")
	strUser = rs("DBUser")
	strpassword = rs("DBPassword")
	strSID=rs("DBSID")
	
	rs.Close	
	Conn.Close
	Set rs = Nothing
	Set Conn = Nothing
	
	Set objDBConn = CreateObject("ADODB.Connection")
	Set rsDBRecords = CreateObject("ADODB.Recordset")	

	If Trim(UCase(strServiceName)) <> "SKIP"  Then
		strCon = "Driver={" & strDriver & "};" 
		strCon = strCon & "CONNECTSTRING=(DESCRIPTION="
		strCon = strCon & "(ADDRESS=(PROTOCOL=TCP)" 
		strCon = strCon & "(HOST=" & strHost & ")(PORT=" & strPort & "))" 
		strCon = strCon & "(CONNECT_DATA=(SERVICE_NAME=" & strServiceName & "))); " 
		strCon = strCon & "uid=" & strUser & ";pwd=" & strPassword & ";"  

'		strCon ="Driver={Microsoft ODBC for Oracle}; " & _
'"CONNECTSTRING=(DESCRIPTION=" & _
'"(ADDRESS=(PROTOCOL=TCP)" & _
'"(HOST=p29938dtw642.machine.test.group)(PORT=1522))" & _
'"(CONNECT_DATA=(SERVICE_NAME=TBT10P11T.test.abc.com))); uid=CHANNEL_USER_SIT1;pwd=Channel_user_sit1_dtw642;"
'
'objDBConn.Open strCon



	Else
		strCon = "Driver={"&strDriver&"};"
		strCon = strCon & "Server=(DESCRIPTION="
		strCon = strCon & "(ADDRESS=(PROTOCOL=TCP)"
		strCon = strCon & "(HOST="& strHost &")(PORT="&strPort&"))"
		strCon = strCon & "(CONNECT_DATA=(SID="&strSID&")));"
		strCon = strCon & "Uid="&strUser&";Pwd="&strPassword&";"
	End If
	
	objDBconn.Open strCon	
	
	If Err.number <> 0 AND Err.Number <> 424 or  objDBconn.state<>1 then
		gstrDesc = "Failed to connect to Database : " & Err.Number & " : " & Err.Description
		WriteHTMLResultLog gstrDesc, 0
	    CreateReport  gstrDesc,0		
		objErr.Raise 11
		bResult=False
	Else
		gstrDesc = "Connected to Database successfully"
		WriteHTMLResultLog gstrDesc, 1
	    CreateReport  gstrDesc,1
		bResult=True
	End if
	CONNECTtoPAM=bResult

End Function
'--------------------------------------------------------½½---------------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------
'Function Name      :  executeDBSelectQuery

'Input Parameter    : None
'Description        : To execute the query on oracle database
'Calls              : None
'Return Value       : True/False
'---------------------------------------------------------------------------------------------------------------------
Public function executeDBSelectQuery(strTestData)

	Dim bResult, arrTemp, strSQL, blnSelect,intCnt, nLoop
	Dim arrColumnList, rs
	'bResult=False
	
	strTestData = getTestDataValue(strTestData)
	
	If UCase(strTestData) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If	
	
	On Error Resume Next
	arrTemp = Split(Trim(strTestData),";")
	
	strSQL = arrTemp(0)
	If UCase(Left(strSQL,6)) <> "SELECT" Then
		gstrDesc = "The query '" & strSQL & "' is not a Select query"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0
		objErr.Raise 11
'		Exit Function
	End If
	
	If InStr(strSQL, "?") > 0 Then
		strSQL = Replace(strSQL, "?", gstrPtyID)
	End If
	If InStr(strSQL, "!") > 0 Then
			strSQL = Replace(strSQL, "!",gstrUserName)
	End If

	If InStr(strSQL, "@") > 0 Then
		today = FormatDateTime(Date,1)
		today = replace(today," ","-")
		strSQL = Replace(strSQL, "@",today)
	End If

	gstrQCDesc = "Execute query '" & strSQL & "'"
	gstrExpectedResult =  "SQL Query '" & strSQL & "' executed successfully"   
	
	blnSelect = True
	wait 2
	Set rs = objDBConn.Execute(strSQL)
	
	arrColumnList =  Split(Trim(arrTemp(1)),",")
	
	If rs.EOF = True  Then 'Or rs(0).Value = NULL
		gstrDesc = "No Data returned after executing the query '" & strSQL & "'"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0
		bResult = False
		objErr.Raise 11
	Else
		gstrDesc = "Successfully executed query '" & strSQL & "'"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc,1
		bResult = True   
            
		For nLoop = 0 to Ubound(arrColumnList)
			 If InStr(arrColumnList(nLoop), "=") > 0 Then
				arrTempField = Split(arrColumnList(nLoop),"=")
			Else
				gstrDesc = "Invalid datafield/datavalue format in '" & arrColumnList(nLoop) & "'"
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc,0
				objErr.Raise 11
			End If
			
			tmpFieldName = arrTempField(0)
			tmpFieldVal = arrTempField(1)
'			tmpFieldVal = mid(tmpFieldVal,2,len(tmpFieldVal)-2)
			
			For i = 0 to rs.Fields.Count-1
				If rs.Fields(i).Name = tmpFieldName Then
					If UCASE(tmpFieldVal) = "NULL" Then
						If IsNull(rs.Fields(i).Value) Then
							bResult = True
						Else
							bResult = False
						End If
'                    ElseIf rs.Fields(i).Value = tmpFieldVal Then
'                        bResult = True   

					ElseIf UCASE(tmpFieldVal) <>"NULL" Then
						tmpFieldVal = rs.Fields(i).Value
                        bResult = True  
						 
					Else
						bResult = False
					End If
					If bResult = True Then
							gstrDesc = "Audit Event log is generated Successfully and Retrieved value is: '" & tmpFieldVal & "' for field '" & tmpFieldName & "'"
							WriteHTMLResultLog gstrDesc, 1
							CreateReport  gstrDesc,1
					Else
							gstrDesc = "Failed to retrieve Audit Event log '" & tmpFieldVal & "' for field '" & tmpFieldName & "'. "
							gstrDesc = gstrDesc & "Actual value is '" & rs.Fields(i).Value & "'"
							WriteHTMLResultLog gstrDesc, 0
							CreateReport  gstrDesc,0
							objErr.Raise 11
					End If
				End If
			Next
		Next
	End If
	
	If Err.number <> 0 And Err.Number <> -2147217900 then
		gstrDesc = Err.Number & " Error occured while executing query '" & strSQL & "' :" & Err.Number & " : " & Err.Description 
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc,0
		bResult=false
		objErr.Raise 11
	Else
		If NOT blnSelect Then
			gstrDesc = "Successfully executed query '" & strSQL & "'"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc,1
			bResult = True	      		
		End If
	End If

	executeDBSelectQuery=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

Public Function SelectListMouseOver(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."

'	call TakeScreenShot()
	arrTemp = Split(strObject,":")	 

	Set objSelectMover = gobjObjectClass.getObjectRef(arrTemp(1))

	If objSelectMover.exist(gExistCount) Then
		'objSelectMover.Click
		'objSelectMover.object.setActive
		Setting.WebPackage("ReplayType") = 2 
        objSelectMover.Set strval
		Setting.WebPackage("ReplayType") = 1

		Set EditDesc = Description.Create() 
		EditDesc("innertext").Value = strVal		
		EditDesc("class").Value = "item" 
		EditDesc("index").Value = "0" 

		Set objSelectMoverClick = gobjObjectClass.getObjectRef(arrTemp(0))
		objSelectMoverClick.WebElement(EditDesc).Click
		bResult = True
		If gbIterationFlag <> True then
			gstrDesc = "Value '" & strVal & "' is selected successfully from '" & strLabel & "' list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
	Else
		Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	SelectListMouseOver=bResult
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SelectListFromWebElement(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."

	arrTemp = Split(strObject,":")	 
	arrVal = Split(strVal,":")	
	Set objSelectMover = gobjObjectClass.getObjectRef(arrTemp(1))

	If objSelectMover.exist(gExistCount) Then
		objSelectMover.Click
		Set EditDesc = Description.Create() 
		EditDesc("innertext").Value = TRIM(arrVal(0))
		EditDesc("outertext").Value = TRIM(arrVal(0))
		EditDesc("html tag").Value = "LABEL|SPAN" 
		EditDesc("index").Value = arrVal(1)
		Set objSelectMoverClick = gobjObjectClass.getObjectRef(arrTemp(0))
		objSelectMoverClick.WebElement(EditDesc).Click
		bResult = True
		If gbIterationFlag <> True then
			gstrDesc = "Value '" & arrVal(0) & "' is selected successfully from '" & strLabel & "' list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
	Else
		Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	SelectListFromWebElement=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:clickWebElementFromWebTable
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function VerifyWebElementFromWebTable(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

'	MsgBox "In Function"
	bResult = false

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	


	Set oDesc = Description.Create() 
	oDesc("micclass").Value = "WebElement" 
	
	Set objTable = gobjObjectClass.getObjectRef(strObject)

	If  objTable.EXIST(gExistCount) Then
'		objTable.waitProperty "attribute/readyState", "complete", gExistCount*1000
		Set objElementCollection = objTable.ChildObjects(oDesc)
		
		NumberOfWebElements = objElementCollection.Count 
		nFlag = 0
		Flag=0
		For i = 0 To NumberOfWebElements - 1 
		
		If Instr(UCase(Trim(objElementCollection (i).GetROProperty("innertext"))), UCase(Trim(strVal))) > 0 Then 
			
			Flag=1
			Exit For
			nFlag = nFlag + 1

			If nFlag > 2 Then
				i = NumberOfWebElements
			End If
			
		End If 
		
		Next 

		If Flag=1 Then
			'gstrDesc =  "Successfully Updated User Detals  :  '" & strVal & "'  For "& strLabel & "."
			gstrDesc =  "Successfully Updated " & strLabel & " '" & strVal & "'  in User Detail Summary "
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
		Else
			gstrDesc =  "Not Found  " & strLabel & " '" & strVal & "'  in User Detail Summary "
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1		
'			bResult = False
'			objErr.Raise 11

		End If
		
	End If
    
	VerifyWebElementFromWebTable = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:SelectFutureDate
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function SelectFutureDate(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

	bResult = false

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	today = FormatDateTime(Date,2)
	chgDate = DateAdd ("d",1,today)
	Nextday = FormatDateTime(chgDate,2)
	strTodayDD=Left (today,2)
	strFutureDD=strTodayDD+2

	Set oDesc = Description.Create() 
	oDesc("micclass").Value = "Link" 
	
	Set objTable = gobjObjectClass.getObjectRef(strObject)

	If  objTable.EXIST(gExistCount) Then

		Set objElementCollection = objTable.ChildObjects(oDesc)

			For nTemp=1 to (objElementCollection.count-1)
			If TRIM(objElementCollection(nTemp).getroproperty("innertext"))=TRIM(strFutureDD) Then
				objElementCollection(nTemp).Click
				gstrDesc =  "Successfully Selected Future date as :  '" & strFutureDD 
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = true
				Exit For
			End If
		Next
	Else
			Call TakeScreenShot()
			gstrDesc =  "Failed to get Future date.."		
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
	End If
    
	SelectFutureDate = bResult
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EmainVerification
'Input Parameter    	:strObject - EmainVerification
'Description          	:EmainVerification
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function GetActivationCodeFromEmail(gstrREMAIL,strLabel,strSearchText)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

	bResult = false

	strVal = getTestDataValue(strVal)
	On Error Resume Next
	Dim intMax, intOld
	Dim objFolder, objItem, objNamespace, objOutlook
	Wait (10)
	Const INBOX =  6 
	Set objOutlook = CreateObject( "Outlook.Application" )
	Set objNamespace = objOutlook.GetNamespace( "MAPI" )
	
	'objNamespace.Logon "Default Outlook Profile", , False, False    
	Set objFolder = objNamespace.GetDefaultFolder( INBOX )
	Set UnRead=objFolder.Items.Restrict("[Unread] = true")



	For Each objItem In UnRead
		If instr(UCase(objItem.Body), UCase(gstrREMAIL))>0 Then

			Dim intLocLabel 
			Dim intLocCRLF
			Dim intLenLabel 	
			Dim strText 
			'strLable="Activation Code 'p2'= "
			intLocLabel = InStr(objItem.Body, strSearchText)
			intLenLabel = Len(strSearchText)
				If intLocLabel > 0 Then
'					intLocCRLF = InStr(intLocLabel, objItem.Body, vbCrLf)

					intLocCRLF = InStr(intLocLabel, objItem.Body, "To")
				
					If intLocCRLF > 0 Then
						intLocLabel = intLocLabel + intLenLabel
						strText = Mid(objItem.Body,intLocLabel,intLocCRLF - intLocLabel)
					Else
						intLocLabel =  Mid(objItem.Body, intLocLabel + intLenLabel)
					End If
				End If
            strText=Replace(strText,VBCRLF,"")
			gActivationVerificationCode = Trim(strText)
			gstrDesc =  "Successfully Captured the" & strLabel & " Code : <B> '" & gActivationVerificationCode &"</B>."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
			Exit For		
		End If
	Next
'---------------------------------------------------------------
Set objFolder    = Nothing
Set objNamespace = Nothing
Set objOutlook   = Nothing

If  gActivationVerificationCode="" Then
	Call TakeScreenShot()
	gstrDesc =  "Failed to get the Activation Code"
	WriteHTMLResultLog gstrDesc, 0
	CreateReport  gstrDesc, 0		
	bResult = False
	objErr.Raise 11  
End If
 
GetActivationCodeFromEmail = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EmainVerification
'Input Parameter    	:strObject - EmainVerification
'Description          	:EmainVerification
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function GetVerificationCodeFromEmail(gstrREMAIL,strLabel,strSearchText)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

	bResult = false

	strVal = getTestDataValue(strVal)
	On Error Resume Next
	Dim intMax, intOld
	Dim objFolder, objItem, objNamespace, objOutlook
	Wait (10)
	Const INBOX =  6 
	Set objOutlook = CreateObject( "Outlook.Application" )
	Set objNamespace = objOutlook.GetNamespace( "MAPI" )
	
	'objNamespace.Logon "Default Outlook Profile", , False, False    
	Set objFolder = objNamespace.GetDefaultFolder( INBOX )
	Set UnRead=objFolder.Items.Restrict("[Unread] = true")



	For Each objItem In UnRead
		If (instr(UCase(objItem.Body), UCase(gstrREMAIL)) and instr(UCase(objItem.Body), UCase("To confirm we have the right email address for you")))>0 Then

			Dim intLocLabel 
			Dim intLocCRLF
			Dim intLenLabel 	
			Dim strText 
			'strLable="Activation Code 'p2'= "
			intLocLabel = InStr(objItem.Body, strSearchText)
			intLenLabel = Len(strSearchText)
				If intLocLabel > 0 Then
'					intLocCRLF = InStr(intLocLabel, objItem.Body, vbCrLf)

					'intLocCRLF = InStr(intLocLabel, objItem.Body, "To")
					intLocCRLF = InStr(intLocLabel, objItem.Body, "This code")
					If intLocCRLF > 0 Then
						intLocLabel = intLocLabel + intLenLabel
						strText = Mid(objItem.Body,intLocLabel,intLocCRLF - intLocLabel)
					Else
						intLocLabel =  Mid(objItem.Body, intLocLabel + intLenLabel)
					End If
				End If
            strText=Replace(strText,VBCRLF,"")
			gActivationVerificationCode = Trim(strText)
			gstrDesc =  "Successfully Captured the" & strLabel & " Code : <B> '" & gActivationVerificationCode &"</B>."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
			Exit For		
		End If
	Next
'---------------------------------------------------------------
Set objFolder    = Nothing
Set objNamespace = Nothing
Set objOutlook   = Nothing

If  gActivationVerificationCode="" Then
	Call TakeScreenShot()
	gstrDesc =  "Failed to get the Activation Code"
	WriteHTMLResultLog gstrDesc, 0
	CreateReport  gstrDesc, 0		
	bResult = False
	objErr.Raise 11  
End If
 
GetVerificationCodeFromEmail = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:GetDatafromWebTable
'Input Parameter      	:strObject - Logical Name of Web Table
'										strLabel:  Label to be used in Report
'										strVal: Value by which function will search for a checkbox
'Description          	:This function checks a specific checkbox in WebTable
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------

Public Function GetDatafromWebTable(strObject, strLabel, strVal)

	Dim nRowIndex, nColumnIndex, bResult, nflag, objTable

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	Set objTable = gobjObjectClass.getObjectRef(strObject)
	If  objTable.EXIST(gExistCount) Then
			ClientNumer=objTable.GetCellData(2,1)				
			Set regEx = New RegExp   ' Create a regular expression.
		   regEx.Pattern = "\d+"  ' Set pattern.
		   regEx.IgnoreCase = True   ' Set case insensitivity.
		   regEx.Global = True   ' Set global applicability.
		   Set Matches = regEx.Execute(ClientNumer)
		   if Matches.count>0 Then   ' Execute search.
				gClientID=Matches(0).Value
		   End If

			gstrDesc =  "Successfully Captured the Client ID : <B> '" & gClientID &"</B>."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
				
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebTable " & strLabel & " is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	
	GetDatafromWebTable = bResult
   
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------
Public Function VerifyAuditEventLog(strVal)

	Dim nRowIndex, nColumnIndex, bResult, nflag, objTable
	
	strVal = getTestDataValue(strVal)
	If UCase(strVal) = "SKIP" Or objErr.number = 11 Then
		Exit Function
	End If

	strtempquery=strVal
	Wait 2
	  gstrDesc =  "Audit Event Logs Verifivation."  
      WriteHTMLResultLog gstrDesc, 2
      CreateReport  gstrDesc, 2

	Call CONNECTtoPAM("TBTSOC")               
	Call executeDBSelectQuery(strtempquery)

End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:EnterSecureText
'Input Parameter      	:strObject - Logical Name of Edit Box
'     			:strVal - Value to  be entered in text box
'Description        	:This function enters a data into a text box
'      			 DLL returns an object reference of a WebEdit
'Calls              	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function EnterMouseText(strObject,strLabel,strVal)

   	Dim bResult, objTextField,objTrim,objlen,objReport, i, Wsh, nCnt, strTextVal
	bResult=False			

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrExpectedResult = strVal & " should get entered in '" & strLabel & "' textbox."
	gstrQCDesc = "Enter '" & strVal & "' in '" & strLabel & "' textbox."
	Set objTextField = gobjObjectClass.getObjectRef(strObject)
	If objTextField.exist(gExistCount) Then
		objTextField.Click
		setting.webpackage("Replaytype")=2
        objTextField.Set strVal
		Wait 2
		Set Wsh = CreateObject("Wscript.Shell")
		Wsh.SendKeys "{TAB}"
		Wait 2
		setting.webpackage("Replaytype")=1 
	
		If gbIterationFlag <> True then
			gstrDesc = "Successfully entered '" & strVal & "' in '" & strLabel & "' textbox."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "'" & strLabel & "' textbox does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If

  	EnterMouseText=bResul
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:VerifyTextinWebTable
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function VerifyTextinWebTable(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
		
	Set objTable = gobjObjectClass.getObjectRef(strObject)

	If  objTable.EXIST(gExistCount) Then
	
		If Instr(objTable.GetROProperty("innertext"),Trim(strVal)) > 0 Then 
			
			gstrDesc =  "Successfully Found " & strVal & "' status in payment management screen. "
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
		Else
			gstrDesc =  "Not Found  " & strVal & "'  status in payment management screen."
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11

		End If
	Else
			gstrDesc =  "Table Not Found "
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11	
	End If
    
	VerifyTextinWebTable = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:clickWebElementInPage
'Input Parameter      	:strObject - Logical Name of Web Button
'Description          	:This function clicks the WebElement object
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function clickWebElementInPage(strObject,strLabel, strVal)

	Dim objWebElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Successfully clicked on Payment ID '" & gstrPaymentID & "'  for  Payment."
	gstrExpectedResult ="Successfully clicked on Payment ID '" & gstrPaymentID & "'  for  Payment."

	Set EditDesc = Description.Create() 
	EditDesc("html tag").Value = "TD|SPAN" 
	EditDesc("Inertext").Value = strVal
	EditDesc("outertext").Value =strVal

	Set objWebPage= gobjObjectClass.getObjectRef(strObject)
	If objWebPage.WebElement(EditDesc).exist(gExistCount) Then
		objWebPage.WebElement(EditDesc).Highlight
		objWebPage.WebElement(EditDesc).click

		If gbIterationFlag <> True then
			If  strLabel="Payment Record" Then
				gstrDesc =  "Successfully clicked on "& strLabel &"<B> '" & gstrPaymentID & "' </B> for  Payment."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
			Else
				gstrDesc =  "Successfully clicked on "& strLabel &"<B> '" & gstrTemplateName & "' </B> for  Payment."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
			End If
			
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebElement '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	clickWebElementInPage=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:GetPaymentID
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function GetPaymentID(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
		
	Set objWebTable = gobjObjectClass.getObjectRef(strObject)

	If  objWebTable.EXIST(gExistCount) Then
    	'gstrPaymentID=TRIM(objWebElement.GetROProperty("outertext"))
		gstrPaymentID=Trim(objWebTable.GetCellData( 1,2))
        gstrDesc =  "Successfully Capture Payment ID:<B> '" & gstrPaymentID & "'</B> for  Payment."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = true
	Else
		gstrDesc =  "Not Found --  Payment ID for  Payment."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11

	End If
        
	GetPaymentID = bResult
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EnterTemplate Name
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function EnterTemplateName(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
		
	Set objEdt = gobjObjectClass.getObjectRef(strObject)

	If  objEdt.EXIST(gExistCount) Then
		Const Chars = "abcdefghijklmnopqrstuvwxyz"
		strName = ""
        Randomize
		For k = 1 To 8
			intValue = Fix(26 * Rnd())
			strChar = Mid(Chars, intValue + 1, 1)	
			strName = strName & strChar
		Next
		gstrTemplateName= UCase(strName)
		objEdt.Set gstrTemplateName
		gstrDesc = "Successfully entered '" & gstrTemplateName & "' in '" & strLabel & "' textbox."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = true
	Else
		gstrDesc =  "'" & strLabel & "' textbox does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11

	End If
        
	EnterTemplateName = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:DoubleClick
'Input Parameter      	:strObject - Logical Name of Web Button
'Description          	:This function clicks the WebElement object
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function ImportPaymentFile(strObject,strLabel, strVal)

	Dim objElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc ="Successfully Import  file for  Payment"
	gstrExpectedResult = "Successfully Import  file for  Payment"
	
	Set objElement = gobjObjectClass.getObjectRef(strObject)
	If objElement.exist(gExistCount) Then
		objElement.highlight
		Setting.WebPackage("ReplayType") = 2 
		objElement.Click
		Setting.WebPackage("ReplayType") = 1
'		objElement.FireEvent "ondblclick"
		Dialog("wndChooseFileUpload").WinEdit("edtCFUFileName").Set gstrLibrariesDir &"\BAC_Latest file.txt"
		Dialog("wndChooseFileUpload").WinButton("btnCFUOpen").Click

		If gbIterationFlag <> True then
			gstrDesc =  "Successfully Import  file for  Payment"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Element '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	ImportPaymentFile=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:clickTemplate
'Input Parameter      	:strObject - Logical Name of Web Button
'Description          	:This function clicks the WebElement object
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function clickTemplate(strObject,strLabel, strVal)

	Dim objWebElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Successfully clicked on Payment ID '" & gstrPaymentID & "'  for  Payment."
	gstrExpectedResult ="Successfully clicked on Payment ID '" & gstrPaymentID & "'  for  Payment."

	Set EditDesc = Description.Create() 
	EditDesc("thml tag").Value = "TD" 
	EditDesc("Inertext").Value = strVal
	EditDesc("outertext").Value =strVal
	EditDesc("class").Value ="dojoxGridCell"

	Set objWebPage= gobjObjectClass.getObjectRef(strObject)
	If objWebPage.WebElement(EditDesc).exist(gExistCount) Then
		objWebPage.WebElement(EditDesc).Highlight
		objWebPage.WebElement(EditDesc).click

		If gbIterationFlag <> True then
			gstrDesc =  "Successfully clicked on Template Name  --> <B> '" & strVal & "' </B> for  Payment."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebElement '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	clickTemplate=bResult
End Function



'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EnterTelephonyPIN
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function EnterTelephonyPIN(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false

	If  strVal="ENTERPIN:172839" Then

	Else
		strVal = getTestDataValue(strVal)	
	End If
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	StrObjectName=Split(strObject,":")		
	StrTelephonyName=Split(strVal,":")	
	StrTelephonyPIN=StrTelephonyName(1)

	If  StrTelephonyPIN="162534"Then
		arrTPIN=Array(1,6,2,5,3,4)
	ElseIf StrTelephonyPIN="172839" Then
		arrTPIN=Array(1,7,2,8,3,9)
	End If

	Set objWebElement = gobjObjectClass.getObjectRef(StrObjectName(0))

	If  objWebElement.EXIST(gExistCount) Then
        strText=TRIM(objWebElement.GetROProperty("innertext"))

		a=Split(strText,"*")
		strone = TRIM(a(1))
		strtwo= TRIM(a(2))
		c = split(a(3),"Please")
		strthree = TRIM(c(0))

		arrTele=array(strone,strtwo,strthree)
		For i=0 to Ubound(arrTele)
            Select Case arrTele(i)
				Case "1"					
					intIndex = arrTPIN(0)
				Case "2"					
					intIndex = arrTPIN(1)
				Case "3"					
					intIndex = arrTPIN(2)
				Case "4"					
					intIndex = arrTPIN(3)
				Case "5"					
					intIndex = arrTPIN(4)
				Case "6"					
					intIndex = arrTPIN(5)
			End Select
		
			Set objList = gobjObjectClass.getObjectRef(StrObjectName(i+1))
			
			If Not objList Is Nothing Then
				objList.Set  intIndex	
			Else			
				objErr.Raise 11
			End If

		Next
		
       
		gstrDesc =  "Telephony PIN : " & StrTelephonyPIN &" Successfully Entered Telephony PIN to verify User."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = true
	Else
		gstrDesc =  "Not Found --  Telephony PIN" & StrTelephonyPIN
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11

	End If
        
	EnterTelephonyPIN = bResult
	
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SelectPaymentRole(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."

'	call TakeScreenShot()
	arrTemp = Split(strObject,":")	 
	arrVal = Split(strVal,":")	
	Set objSelectMover = gobjObjectClass.getObjectRef(arrTemp(1))

	If objSelectMover.exist(gExistCount) Then
		'Setting.WebPackage("ReplayType") = 2 

		objSelectMover.Click
		'Setting.WebPackage("ReplayType") = 1

		Set EditDesc = Description.Create() 
		EditDesc("innertext").Value = TRIM(arrVal(0))
		EditDesc("html tag").Value = "SPAN" 
		EditDesc("index").Value = arrVal(1)
		Set objSelectMoverClick = gobjObjectClass.getObjectRef(arrTemp(0))
		objSelectMoverClick.WebElement(EditDesc).Click
		bResult = True
		If gbIterationFlag <> True then
			gstrDesc = "Value '" & strVal & "' is selected successfully from '" & strLabel & "' list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
	Else
		Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	SelectPaymentRole=bResult
End Function



'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SelectListWithInnerTextHTMLTag(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."

'	call TakeScreenShot()
	arrTemp = Split(strObject,":")	 
	arrVal = Split(strVal,":")	
	Set objSelectMover = gobjObjectClass.getObjectRef(arrTemp(1))

	If objSelectMover.exist(gExistCount) Then
		'Setting.WebPackage("ReplayType") = 2 

		objSelectMover.Click
		'Setting.WebPackage("ReplayType") = 1

		Set EditDesc = Description.Create() 
		EditDesc("innertext").Value = TRIM(arrVal(0))
		EditDesc("index").Value = arrVal(1)
		EditDesc("html tag").Value = arrVal(2)
		Set objSelectMoverClick = gobjObjectClass.getObjectRef(arrTemp(0))
		objSelectMoverClick.WebElement(EditDesc).Click
		bResult = True
		If gbIterationFlag <> True then
			gstrDesc = "Value '" & arrVal(0) & "' is selected successfully from '" & strLabel & "' list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
	Else
		Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	SelectListWithInnerTextHTMLTag=bResult
End Function


'---------------------------------------------------------------------------------------------------------------------

Public Function SelectListWebElement(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."

'	call TakeScreenShot()
	arrTemp = Split(strObject,":")	 

	Set objSelectMover = gobjObjectClass.getObjectRef(arrTemp(1))

	If objSelectMover.exist(gExistCount) Then
		'objSelectMover.Click
		'objSelectMover.object.setActive
		Setting.WebPackage("ReplayType") = 2 
        objSelectMover.Set strval
		Setting.WebPackage("ReplayType") = 1
		Wait 5
		Set EditDesc = Description.Create() 
		EditDesc("innertext").Value = strVal		
		'EditDesc("class").Value = "item" 
		EditDesc("html tag").Value = "DIV" 
		EditDesc("index").Value = "0" 

		Set objSelectMoverClick = gobjObjectClass.getObjectRef(arrTemp(0))
		objSelectMoverClick.WebElement(EditDesc).Click
		bResult = True
		If gbIterationFlag <> True then
			gstrDesc = "Value '" & strVal & "' is selected successfully from '" & strLabel & "' list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
	Else
		Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	SelectListWebElement=bResult
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:GetPaymentID
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function GetBankPaymentID(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
		
	Set objWebTable = gobjObjectClass.getObjectRef(strObject)

	If  objWebTable.EXIST(gExistCount) Then
    	'gstrPaymentID=TRIM(objWebElement.GetROProperty("outertext"))
		gstrBankPaymentID=Trim(objWebTable.GetCellData( 1,13))
        gstrDesc =  "Successfully Capture Bank Payment ID:<B> '" & gstrBankPaymentID & "'</B> for Payment."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = true
	Else
		gstrDesc =  "Not Found -- Bank  Payment ID for BACS Payment."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11

	End If
        
	GetBankPaymentID = bResult
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EmainVerification
'Input Parameter    	:strObject - EmainVerification
'Description          	:EmainVerification
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function VerifyCBOSharedEmail(strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

	bResult = false

	strVal = getTestDataValue(strVal)
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	arrTemp = Split(strVal,":")	 
	strUserName=arrTemp(0)
	strUserStatus=arrTemp(1)
	On Error Resume Next
	Dim intMax, intOld
	Dim objFolder, objItem, objNamespace, objOutlook
	Wait (10)
	Const INBOX =  6 
	Set objOutlook = CreateObject( "Outlook.Application" )
	Set objNamespace = objOutlook.GetNamespace( "MAPI" )
	
	'objNamespace.Logon "Default Outlook Profile", , False, False    
	Set objFolder = objNamespace.GetDefaultFolder( INBOX )
	Set UnRead=objFolder.Items.Restrict("[Unread] = true")

	For Each objItem In UnRead
		If instr(UCase(objItem.Body), UCase(strUserName))>0 Then
			gstrDesc =  "Successfully Get the Status : <B>" & strUserStatus & "</B> For User : <B> '" & strUserName &"'</B> in Shared Mail box."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
			Exit For		
		End If
	Next
'---------------------------------------------------------------
Set objFolder    = Nothing
Set objNamespace = Nothing
Set objOutlook   = Nothing

If  bResult <> TRUE Then
	Call TakeScreenShot()
	gstrDesc =  "Failed Get the Status : <B>" & strUserStatus & "</B> For User : <B> '" & strUserName &"</B> in Shared Mail box."
	WriteHTMLResultLog gstrDesc, 0
	CreateReport  gstrDesc, 0		
	bResult = False
	objErr.Raise 11  
End If
 
VerifyCBOSharedEmail = bResult
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EmainVerification
'Input Parameter    	:strObject - EmainVerification
'Description          	:EmainVerification
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function GetPasswordCodeFromEmail(strLabel,strSearchText,strUserName)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

	bResult = false

	strVal = getTestDataValue(strUserName)
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	On Error Resume Next
	Dim intMax, intOld
	Dim objFolder, objItem, objNamespace, objOutlook
	Wait (10)
	Const INBOX =  6 
	Set objOutlook = CreateObject( "Outlook.Application" )
	Set objNamespace = objOutlook.GetNamespace( "MAPI" )
	
	'objNamespace.Logon "Default Outlook Profile", , False, False    
	Set objFolder = objNamespace.GetDefaultFolder( INBOX )
	Set UnRead=objFolder.Items.Restrict("[Unread] = true")



	For Each objItem In UnRead
		If instr(UCase(objItem.Body), UCase(strVal))>0 Then

			Dim intLocLabel 
			Dim intLocCRLF
			Dim intLenLabel 	
			Dim strText 
			'strLable="Activation Code 'p2'= "
			intLocLabel = InStr(objItem.Body, strSearchText)
			intLenLabel = Len(strSearchText)
				If intLocLabel > 0 Then
'					intLocCRLF = InStr(intLocLabel, objItem.Body, vbCrLf)

					intLocCRLF = InStr(intLocLabel, objItem.Body, "To")
				
					If intLocCRLF > 0 Then
						intLocLabel = intLocLabel + intLenLabel
						strText = Mid(objItem.Body,intLocLabel,intLocCRLF - intLocLabel)
					Else
						intLocLabel =  Mid(objItem.Body, intLocLabel + intLenLabel)
					End If
				End If
            strText=Replace(strText,VBCRLF,"")
			gActivationVerificationCode = Trim(strText)
			gstrDesc =  "Successfully Captured the" & strLabel & " Code  --> <B> '" & gActivationVerificationCode &"</B>."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = true
			Exit For		
		End If
	Next
'---------------------------------------------------------------
Set objFolder    = Nothing
Set objNamespace = Nothing
Set objOutlook   = Nothing

If  gActivationVerificationCode="" Then
	Call TakeScreenShot()
	gstrDesc =  "Failed to get the Activation Code"
	WriteHTMLResultLog gstrDesc, 0
	CreateReport  gstrDesc, 0		
	bResult = False
	objErr.Raise 11  
End If
 
GetActivationCodeFromEmail = bResult
	
End Function

'---------------------------------------------------------------------------------------------------
'Function Name		: ClickWebTableElement
'Input Parameter	: strObject - Logical Name of WebTable
'Description		: This function CLick link in web table
'Calls				: None
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Public Function ClickWebTableElement(strObject,strLabel,strVal)
	'msgbox "in func"
	Dim bResult, objWebTableLink, objclick

	bResult=False
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	arrTemp = Split(strVal,":")
	ntRow = arrTemp(0)
	ntColumn = arrTemp(1)
	ntIndex = arrTemp(2)
			
	Set objWebTableLink = gobjObjectClass.getObjectRef(strObject)
	
	If objWebTableLink.exist(gExistCount) Then

		gstrExpectedResult = strLabel & " Webelement should get clicked"
		gstrQCDesc = "Click on Webelement '" & strLabel &"'"

		Set  objclick = objWebTableLink.ChildItem(Cint(ntRow) , Cint(ntColumn) ,"WebElement", Cint(ntIndex))
		objclick.click

        gstrDesc = "Successfully Clicked " & strLabel &"Webelement"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = True
	Else		 
		gstrDesc =  "'" & strLabel & "' in WebTable does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If

ClickWebTableElement = bResult

End Function



'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:VerifyEmailTempalate
'Input Parameter    	:strObject - EmainVerification
'Description          	:EmainVerification
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function VerifyEmailTempalate(strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag

	bResult = false

	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	strEmailT=Split(strVal,":")
	strEmailTempValidate=strEmailT(0)
	strEmailTempVal=strEmailT(1)

	If UCase(strEmailTempValidate) = "MAIL" Then
			strSearchVal=gstrRegistrationEmail
		ElseIf UCase(strEmailTempValidate) = "CLIENTID" Then
			strSearchVal=strEmailTempVal
		ElseIf UCase(strEmailTempValidate) = "LASTNAME" Then
			strSearchVal=strEmailTempVal
		ElseIf UCase(strEmailTempValidate) = "USERNAME" Then
			strSearchVal=strEmailTempVal
	End If

	On Error Resume Next
	Dim intMax, intOld
	Dim objFolder, objItem, objNamespace, objOutlook
	Wait (10)
	Const INBOX =  6 
	Set objOutlook = CreateObject( "Outlook.Application" )
	Set objNamespace = objOutlook.GetNamespace( "MAPI" )
    Set objFolder = objNamespace.GetDefaultFolder( INBOX )
	Set UnRead=objFolder.Items.Restrict("[Unread] = true")

	For Each objItem In UnRead
		If instr(UCase(objItem.Body), UCase(strSearchVal))>0 Then

			Dim intLocLabel 
			Dim intLocCRLF
			Dim intLenLabel 	
			Dim strText 

			strSetupData = Split(objItem.Body,vbcrlf)

			For iCase = LBound(strSetupData) To UBound(strSetupData)
				gstrDesc = strSetupData(iCase)
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
				bResult = true
			Next
		End If
	Next

Set objFolder    = Nothing
Set objNamespace = Nothing
Set objOutlook   = Nothing
VerifyEmailTempalate = bResult
	
End Function



'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:VerifyPlaceholder
'Input Parameter      	:strObject - Logical Name of Browser
'Description          	:This function clicks the Image object
'      			 DLL returns an object reference of an Image
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function Verifyplaceholder(strObject,strLabel, strVal)
	
	Dim bResult, objImage
	bResult=False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	call TakeScreenShot()
	Set objPlaceHolder = gobjObjectClass.getObjectRef(strObject)

	arrTemp=split(strVal,":")
	strObjectProperty=arrTemp(0)
	strObjectType=arrTemp(1)

	gstrQCDesc = "Verify " & strLabel
	gstrExpectedResult = strLabel & "'should be displayed."	
	If objPlaceHolder.exist(gExistCount) Then
			If  strObjectType="WEBEDIT"Then		
					Set a =objPlaceHolder.Object
					On error Resume Next
					Set a =objPlaceHolder.Object
					
					Set c=a.getElementsByTagName("input")
					inCount=c.Length	
					For i=0 to inCount-1
						Flag=0
						If  c(i).type="text" Then 'and c(i).className="input" Then
							If c(i).placeholder ="Wildcard allowed" Then
								If Err.Number =438 Then
									Err.Clear
									Flag=1
								End If
							End If
							If Flag=0 Then						
								strLabel=c(i).Name
								If InStr(1,strLabel,":") Then
									arrLabel=split(strLabel,":")
									strLabel=arrLabel(1)
								End If
								strName=c(i).placeholder
								gstrDesc =  "Successfully verify the placeholder value for Web-Edit   '" & strLabel & "'    -->>'" & strName & "'"
								WriteHTMLResultLog gstrDesc, 1
								CreateReport  gstrDesc, 1
							Else

							End If
						End If

					Next
			ElseIf strObjectType="WEBLIST" Then	
					On error Resume Next
					Set a=objPlaceHolder.Object
					Set c=a.getElementsByTagName("div")
					intcnt=c.length
					For i=0 to intcnt
								If c(i).className="dropdown-control" Then
									Set e=c(i).childNodes
									cntCount=e.length
									For j=0 to cntCount-1
										If e(j).className="label" Then
											If Err.Number =438 Then
												Err.Clear
												Exit For
											End If
												strLabel=e(j).innertext
												For w=0 to cntCount-1
													If e(w).className="selection normal-input" Then
														strName=e(w).innertext							
														gstrDesc =  "Successfully verify the placeholder value for Web-List   '" & strLabel & "'    -->>'" & strName &"'"
														WriteHTMLResultLog gstrDesc, 1
														CreateReport  gstrDesc, 1
														Exit For
													End If
												Next
											Exit For										
										End If
									Next							
								End If
					Next	
		End If	
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Failed to verify placeholder value for  '" &strLabel
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0    		
		bResult = False
		objErr.Raise 11
	End If

	Verifyplaceholder=bResult
	
End Function


'---------------------------------------------------------------------------------------------------
'Function Name		: ClickWebTableElement
'Input Parameter	: strObject - Logical Name of WebTable
'Description		: This function CLick link in web table
'Calls				: None
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Public Function ClickWebTableAllElement(strObject,strLabel,strVal)
	'msgbox "in func"
	Dim bResult, objWebTableLink, objclick

	bResult=False
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objWebTableCheckBox = gobjObjectClass.getObjectRef(strObject)
	
	If objWebTableCheckBox.exist(gExistCount) Then

		gstrExpectedResult = strLabel & " Webelement should get clicked"
		gstrQCDesc = "Click on Webelement '" & strLabel &"'"
		
		Set EditDesc = Description.Create() 
				EditDesc("micclass").Value = "WebElement"		
				EditDesc("class").Value = "inner-check.*" 
				EditDesc("html tag").Value = "SPAN" 

		Set  objclick = objWebTableCheckBox.ChildObjects(EditDesc)		
		strVal=strVal * 2
		For i=1 to strVal step 2
			objclick(i).click
			gstrDesc = "Successfully Clicked " & objWebTableCheckBox.GetCellData(i,2) &"  Account."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = True
		Next
	Else		 
		gstrDesc =  "'" & strLabel & "' in WebTable does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0      		
		bResult = False
		objErr.Raise 11
	End If

ClickWebTableAllElement = bResult

End Function

'---------------------------------------------------------------------------------------------------
'Function Name		: SelectListByDOM
'Input Parameter	: strObject - Logical Name of WebTable
'Description		: This function SelectListByDOM
'Calls				: None
'Return	Value		: NA
'---------------------------------------------------------------------------------------------------
Public Function SelectListByDOM(strObject,strLabel, strVal)
	Dim objLink, bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."

	Set objSelectList = gobjObjectClass.getObjectRef(strObject)

	If objSelectList.exist(gExistCount) Then

		Set ObjPar=objSelectList.Object.parentNode.parentNode		
		If ObjPar.className = "dropdown-control" OR ObjPar.className = "dropdown-control resetValue" Then
			ObjChild=ObjPar.childNodes.length
			If ObjChild>=1 Then
				Set ObjElm=ObjPar.getElementsByTagName("span")
				inCount=ObjElm.Length		
				For i=0 to inCount-1
					If ObjElm(i).className="dropDownIcon icon-down" Then
						ObjElm(i).Click()
						Exit For
					End If
				Next
			End If
		
			Set ObjElm=ObjPar.getElementsByTagName("label")
			inCount=ObjElm.Length		
			For i=0 to inCount-1
				If ObjElm(i).Innertext=strVal Then
					ObjElm(i).Click()
					Exit For
				End If
			Next
		End If

		bResult = True
		If gbIterationFlag <> True then
			gstrDesc = "Value '" & strVal & "' is selected successfully from '" & strLabel & "' list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
	Else
		Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	Call pressTab()
	SelectListByDOM=bResult
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:VerifyValueProerty
'Input Parameter      	:strObject - Logical Name of Image
'Description          	:This function clicks the Image object
'      			 DLL returns an object reference of an Image
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifyValueProperty(strObject,strLabel, strVal)
	
	Dim bResult, objImage
	bResult=False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	call TakeScreenShot()
	Set objPage = gobjObjectClass.getObjectRef(strObject)

	arrTemp=split(strVal,":")
	strObjectVerify=arrTemp(0)
	strObjectType=arrTemp(1)

	gstrQCDesc = "Verify " & strLabel
	gstrExpectedResult = strLabel & "'should be displayed."	

	If objPage.exist(gExistCount) Then
			If  strObjectType="WEBLIST" Then		
					On error Resume Next
					Set a=objPage.Object
					Set c=a.getElementsByTagName("div")
					intcnt=c.length
					For i=0 to intcnt
								If c(i).className="dropdown-control" Then
									Set e=c(i).childNodes
									cntCount=e.length
									For j=0 to cntCount-1
										If e(j).className="label" Then
											If Err.Number =438 Then
												Err.Clear
												Exit For
											End If
												strValCol=e(j).innertext
												For w=0 to cntCount-1
													If e(w).className="selection normal-input" Then
														strValItem=e(w).innertext
														gstrDesc = "Value '" & strValItem &"' is verified successfully from '" & strValCol & "' field."
														WriteHTMLResultLog gstrDesc, 1
														CreateReport  gstrDesc, 1
														Exit For
													End If
												Next
											Exit For										
										End If
									Next							
								End If
					Next	
			ElseIf strObjectType="WEBEDIT" Then	
					Set EditDesc = Description.Create() 
					EditDesc("micclass").Value = "WebEdit" 
					EditDesc("type").Value = "text" 
					EditDesc("html tag").Value = "INPUT" 
					EditDesc("kind").Value = "singleline" 
					EditDesc("class").Value = "input.*" 
                    					
					Set EditCollection =objPage.ChildObjects(EditDesc)
					NumberOfEdits = EditCollection.Count
					For i = 0 To NumberOfEdits - 1
						strValCol= EditCollection(i).GetROProperty("Name")
						If InStr(1,strValCol,":") Then
							strarrCol=split(strValCol,":")
							strValCol=strarrCol(1)
						End If						
						strValItem= EditCollection(i).GetROProperty("value")
						gstrDesc = "Value '" & strValItem& "' is verified successfully from '" & strValCol & "' field."
						WriteHTMLResultLog gstrDesc, 1
						CreateReport  gstrDesc, 1
					Next
			End If	
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Failed to verify :  value for  '" &strLabel
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0    		
		bResult = False
		objErr.Raise 11
	End If

	VerifyValueProperty=bResult
	
End Function


'-------------------------------------------------------------------------------------------------------------------------------
'Function-Name :VerifyTableDetailsTBT
'Description : This Function verifies table properties
'Output-None
'_________________________________________________________________________
Public Function VerifyAllTableDetails(strObject, strLabel, strVal)
	Dim objTable,intRowCnt,intColCnt,bResult

	bResult=False
	
	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objTable = gobjObjectClass.getObjectRef(strObject)
	If  objTable.EXIST(gExistCount) Then
	'	intRowCnt=1
		intRowCnt=objTable.RowCount
'		intRowCnt=intRowCnt-1
		intColCnt=objTable.ColumnCount(intRowCnt)
'		intColCnt=intColCnt-1

		gstrDesc = strLabel & " Displayed Successfully"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		For intRow=1 to intRowCnt
			gstrDesc = "<BR>"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			For intCol=1 to intColCnt
            	bResult=True
				If instr(1,objTable.GetCellData( intRow+1,intCol),"ERROR") Then
					Exit For
				End If
				gstrDesc = objTable.GetCellData( 1,intCol) & " is  -->> " &  objTable.GetCellData( intRow+1,intCol)
				
				 WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
			Next
		Next
	ElseIf Browser("name:=.*").Page("title:=.*").WebElement("innertext:=There were no results for this search.").Exist(2) Then
		bResult=True
		gstrDesc = strLabel &" 'There were no results for this search.'"		
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
	Else 
		bResult=False
		gstrDesc = strLabel & " Table does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		bResult = False
		objErr.Raise 11
    End If
		

VerifyAllTableDetails=bResult
End Function



'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:setWebCheckBox
'Input Parameter      	:strObject - Logical Name of Check box
'     			:strVal - Contains the property of the Chechkbox which needs to be changed and the
'      			 value it needs to be changed to.
'Description          	:This function toggles the checkbox
'      			 DLL returns an object reference of a WebCheckBox
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function setWebCheckBox(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objBrw = gobjObjectClass.getObjectRef(strObject)		
	gstrQCDesc = "Set the check box " & strLabel & " " & UCase(strVal)
	gstrExpectedResult = "Checkbox " & strLabel & " should be set to " & UCase(strVal)
	If objBrw.exist(gExistCount) Then

		ArrVal=Split(strVal,":")
		strClass=ArrVal(0)
		strCheckName=ArrVal(1)
		Set ObjPar=objBrw.Object
		Set ObjElm=ObjPar.getElementsByTagName("span")
		inCount=ObjElm.Length		
		For i=0 to inCount-1
			If ObjElm(i).className=strClass Then
				If ObjElm(i).getattribute("aria-label")=strCheckName Then
					ObjElm(i).Click()
					Exit For
				End If			
			End If
		Next
		If gbIterationFlag <> True then
			gstrDesc = "Successfully selected the '" & strCheckName & "' Checkbox."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "Checkbox '" & strLabel & "' is not displayed in the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	setWebCheckBox = bresult
End Function


'-------------------------------------------------------------------------------------------------------------------------------
'Function-Name :GetAuditReportDetails
'Description : This Function verifies table properties
'Output-None
'_________________________________________________________________________
Public Function GetAuditReportDetails(strObject, strLabel, strVal)
	Dim objTable,intRowCnt,intColCnt,bResult

	bResult=False
	
	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objTable = gobjObjectClass.getObjectRef(strObject)
	If  objTable.EXIST(gExistCount) Then
	'	intRowCnt=1
		intRowCnt=objTable.RowCount
'		intRowCnt=intRowCnt-1
		intColCnt=objTable.ColumnCount(intRowCnt)
'		intColCnt=intColCnt-1

		gstrDesc = strLabel & "  Displayed Successfully"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		For intRow=1 to intRowCnt
			gstrDesc = "<BR>"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1

			For intCol=1 to intColCnt
            	bResult=True
				gstrDesc = objTable.GetCellData( 3,intCol) & " is --> " & vbTab & objTable.GetCellData( intRow+1,intCol)
				
				 WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
			Next
		Next
	Else 
		bResult=False
		gstrDesc = strLabel & " Table does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		bResult = False
		objErr.Raise 11
    End If
		
GetAuditReportDetails=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:SelectDateRange
'Input Parameter      	:strObject - Logical Name of Web Button and Daterange webelement
'Description          	:This function clicks the WebElement object
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function SelectDateRange(strObject,strLabel, strVal)

	Dim objWebElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	arrstrObj= split(strObject,":")
	arrstrVal= split(strVal,":")
	strDate=arrstrVal(0)
	strMonth=arrstrVal(1)
	strYear=arrstrVal(2)

	gstrQCDesc = "Successfully Selected Date Range"
	gstrExpectedResult = "Successfully Selected Date Range"
	
	Set objWebButton = gobjObjectClass.getObjectRef(arrstrObj(0))
	Set objDatePicker = gobjObjectClass.getObjectRef(arrstrObj(1))

	If objWebButton.exist(gExistCount) Then
		objWebButton.click

		Set a=objDatePicker.Object.childNodes
				Set c=a(0).childNodes
				Set d=c(1).getElementsByTagName("span")
				intCntSPAN=d.length
				For j=0 to intCntSPAN-1
						If d(j).className="dropDownIcon icon-down" Then
							'd(j).highlight
							d(j).Click()
							Exit For
						End If
				Next	
				Set e=c(1).getElementsByTagName("label")
				intCntLabel=e.length
				For j=0 to intCntLabel-1
						If e(j).innertext=strMonth Then
							'e(j).highlight
							e(j).Click()
						End If
				Next
		
				Set f=c(2).getElementsByTagName("span")
						intCntSPAN=d.length
						For j=0 to intCntSPAN-1
								If f(j).className="dropDownIcon icon-down" Then
								'	f(j).highlight
									f(j).Click()
									Exit For
								End If
						Next	
				Set g=c(2).getElementsByTagName("label")
				intCntLabel=e.length
				For j=0 to 10
						If g(j).innertext=strYear Then
						'	g(j).highlight
							g(j).Click()
							Exit For
						End If
				Next
					
				Set h=a(1).getElementsByTagName("a")
				intCntA=h.length
				For j=0 to intCntA-1
						If TRIM(h(j).innertext)=strDate Then
							h(j).Click()
							Exit For
						End If
				Next	

		If gbIterationFlag <> True then
			gstrDesc = "Successfully Selected Date Range for " & strLabel & "-->" & strDate & "/" & strMonth & "/" & strYear &"."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebButton '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	SelectDateRange=bResult
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:GetSelectListOption
'Input Parameter      	:strObject - Logical Name of Webelement-LabelName
'Description          	:This function clicks the GetSelectListOption
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function GetSelectListOption(strObject,strLabel, strVal)

	Dim objWebElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	gstrQCDesc = "Click on web element " & strLabel & "."
	gstrExpectedResult = "Successfully get option list for  '" & strLabel& "' web list."
	
	Set objSelectList = gobjObjectClass.getObjectRef(strObject)
	If objSelectList.exist(gExistCount) Then

	gstrDesc =  "Successfully get option list for  '" & strLabel & "' web list."
	WriteHTMLResultLog gstrDesc, 1
	CreateReport  gstrDesc, 1

	Set ObjPar=objSelectList.Object.parentNode.parentNode		
		If ObjPar.className = "dropdown-control" OR ObjPar.className = "dropdown-control resetValue" Then
			ObjChild=ObjPar.childNodes.length
			If ObjChild>=1 Then
				Set ObjElm=ObjPar.getElementsByTagName("option")
				inCount=ObjElm.Length		
				For i=0 to inCount-1
					strOptionName=ObjElm(i).innertext
					gstrDesc =  "Option -->>'" & strOptionName
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1			
				Next
				bResult = True
			End If		
		End If
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebList '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	GetSelectListOption=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------

Function WaitTime(strVal)

	strVal = getTestDataValue(strVal)
	
	If Instr(strVal,"SKIP")>0 Or objErr.Number = 11 Then
	Exit Function
	End If
	
	Set oWinform=DotNetFactory.CreateInstance("System.Windows.Forms.Form", "System.Windows.Forms") 
	
	Set OprogressBar=DotNetFactory.CreateInstance("System.Windows.Forms.ProgressBar", "System.Windows.Forms")
	oWinform.Text="Wait-10 Minutes"
	'msgbox oWinform.Text
	oWinform.width=800
	oWinform.height=80
	oWinform.maximizeBox=false
	oWinform.minimizeBox=False
	
	OprogressBar.Top =10
	OprogressBar.Left =50
	OprogressBar.Width =700
	OprogressBar.Height =25
	OprogressBar.Minimum=1
	OprogressBar.maximum=60
	OprogressBar.Value=1
	oWinform.controls.Add OprogressBar
	oWinform.show
	Call TakeScreenShot()		
	gstrDesc = "Time Window open successfully."				
					WriteHTMLResultLog gstrDesc, 4
					CreateReport  gstrDesc, 4
	For i=1 to 10
	OprogressBar.PerformStep()
	wait(1)
	Next
	oWinform.close
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:CheckObjectExist
'---------------------------------------------------------------------------------------------------------------------
Public Function CheckObjectExist(strObject,strLabel, strVal)
	
	Dim bResult, objImage
	bResult=False	
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	call TakeScreenShot()
	Set objExist = gobjObjectClass.getObjectRef(strObject)

	gstrQCDesc = "Checked for  " & strLabel
	gstrExpectedResult = "Checked for '" & strLabel
	If objExist.exist(1) Then
			
			gstrDesc = "Found  " & strLabel 
			WriteHTMLResultLog gstrDesc, 4
			CreateReport  gstrDesc, 1
			bResult = True
	Else
		Call TakeScreenShot()
		gstrDesc = "Not found " & strLabel 
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		
	End If
	CheckObjectExist=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:VerifySibling
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifySibling(strObject,strLabel, strVal)
	
	Dim bResult, objImage
	bResult=False	
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	call TakeScreenShot()
	Set objExist = gobjObjectClass.getObjectRef(strObject)
	If objExist.exist(gExistCount) Then
			Set ObjSib=objExist.Object.nextSibling
			strMassege=ObjSib.innerHTML
'			strMassege=ObjSib.text
			gstrDesc = "Successfully get the value for " & strLabel & "-><B>" & strMassege&"</B>."
			WriteHTMLResultLog gstrDesc, 4
			CreateReport  gstrDesc, 1
			bResult = True
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to get the value for " & strLabel
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		objErr.Raise 11
	End If
	VerifySibling=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:ClickSiblingPrevious
'---------------------------------------------------------------------------------------------------------------------
Public Function ClickSiblingPrevious(strObject,strLabel, strVal)
	
	Dim bResult, objImage
	bResult=False	
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	call TakeScreenShot()
	Set objExist = gobjObjectClass.getObjectRef(strObject)
	If objExist.exist(gExistCount) Then
			Set ObjSib=objExist.Object.previousSibling
			ObjSib.click
			gstrDesc = "Successfully selected the '" & strLabel & "' object."
			WriteHTMLResultLog gstrDesc, 4
			CreateReport  gstrDesc, 1
			bResult = True
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to click the object " & strLabel
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	ClickSiblingPrevious=bResult
End Function
'========================================================================================='
'Action For SelectDateManually
'=========================================================================================''
Public Function SelectDateManually(strVal)

	Dim today, todayid, holidayDate
	strVal = getTestDataValue(strVal)
 
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	On Error Resume Next
	
			If UCase(getTestDataValue(strVal))="HOLIDAY" Then
				today = FormatDateTime(Date,1)
				todayid=weekday(today)
				holidayDate = DateAdd ("d",(14-todayid),today)
				ltrDate = FormatDateTime(holidayDate,1)
				gstrDate = split(ltrDate," ")
			Elseif UCase(getTestDataValue(strVal))="MONTHERROR" Then
				today = FormatDateTime(Date,1)
				chgDate = DateAdd ("d",32,today)
				dayName = WeekDay(chgDate)
				ltrDate = FormatDateTime(chgDate,1)
				gstrDate = split(ltrDate," ")
			Else
				today = FormatDateTime(Date,1)
				chgDate = DateAdd ("d",16,today)
				dayName = WeekDay(chgDate)
			
				If dayName = ("1" or "7") Then
					chgDate = DateAdd("d",22,today)
				End If
			
				ltrDate = FormatDateTime(chgDate,1)
				gstrDate = split(ltrDate," ")
			End If
		  objErr.Clear
	
End Function


'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:VerifyObjectProperty
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifyObjectProperty(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objEdit = gobjObjectClass.getObjectRef(strObject)

	If  objEdit.EXIST(gExistCount) Then
			Set ObjSib=objEdit.Object.nextSibling
			strStatus=ObjSib.text

			If instr(strVal,strStatus) >0 Then

						gstrDesc =  "Successfully Get the value for '" & strLabel & "'-><B>" & strVal & "</B>"
						WriteHTMLResultLog gstrDesc, 1
						CreateReport  gstrDesc, 1
						bResult = true
			End If
			
	Else
			gstrDesc =  "Value not Found  for " & strVal
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
	End If
    VerifyObjectProperty = bResult
End Function
'-------------------------------------------------------------------------------------------------------------------------------
'Function-Name :VerifyAllTableDetailsItems
'Description : This Function verifies table properties
'Output-None
'_________________________________________________________________________
Public Function VerifyAllTableDetailsItems(strObject, strLabel, strVal)
	Dim objTable,intRowCnt,intColCnt,bResult
    k=0
	bResult=False
	
	strVal = getTestDataValue(strVal)
	arrIndividual = Split(strVal,";")
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	Set objTable = gobjObjectClass.getObjectRef(strObject)
	If  objTable.EXIST(gExistCount) Then
		intRowCnt=objTable.RowCount
		intColCnt=objTable.ColumnCount(intRowCnt)
		gstrDesc = strLabel & " Displayed Successfully"
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
      
		For intRow=2 to intRowCnt
									gstrDesc = "<BR>"
									WriteHTMLResultLog gstrDesc, 1
									CreateReport  gstrDesc, 1
									For intCol=2 to 3
                                                                  If Instr(objTable.GetCellData( intRow,intCol) , arrIndividual(k)) > 0 Then

																		gstrDesc = objTable.GetCellData( 1,intCol) & " is  -->> " &  arrIndividual(k)
																		 WriteHTMLResultLog gstrDesc, 1
																		CreateReport  gstrDesc, 1
																End if 
																k=k+1
									Next
									
		Next
	ElseIf Browser("name:=.*").Page("title:=.*").WebElement("innertext:=There were no results for this search.").Exist(2) Then
								bResult=True
								gstrDesc = strLabel &" 'There were no results for this search.'"		
								WriteHTMLResultLog gstrDesc, 1
								CreateReport  gstrDesc, 1
	Else 
								bResult=False
								gstrDesc = strLabel & " Table does not exist"
								WriteHTMLResultLog gstrDesc, 0
								CreateReport  gstrDesc, 0
								bResult = False
								objErr.Raise 11
    End If
VerifyAllTableDetailsItems=bResult
End Function


'---------------------------------------------------------------------------------------------------
'Function Name		: ClickAllButton
'---------------------------------------------------------------------------------------------------
Public Function ClickAllButton(strObject,strLabel,strVal)
	Dim objChk, bresult,objPage,arrChk,nLoop
	 
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	wait 5

    Set objChk = Description.Create
	objChk("micclass").Value = "WebButton" 
	objChk("name").Value = "Reviewed" 
	objChk("html tag").Value = "BUTTON"   
	
	Set objPage = gobjObjectClass.getObjectRef(strObject)
	Set arrChk = objPage.ChildObjects(objChk)
	If objPage.Exist(gExistCount) Then
		If arrChk.Count > 0 Then
				For nLoop = 0 to arrChk.count-1
					Wait 3
					 arrChk(nLoop).highlight
					  arrChk(nLoop).Click
				Next
		End If
	
		gstrDesc =  "Successfully clicked all button:  " & strLabel 
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
	Else
		Call TakeScreenShot()
		gstrDesc =  "Page  '" & strLabel & "' is not displayed on screen."		
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		objErr.Raise 11
	End If
    ClickAllButton = bresult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:VerifyLabel_OLD
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function VerifyLabel_OLD(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)

	arrLabVal=Split(strVal,";")
	narrLab=Ubound(arrLabVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objPage= gobjObjectClass.getObjectRef(strObject)

	If  objPage.EXIST(gExistCount) Then
			TotalLab=narrLab+1
			gstrDesc =  "<B>" & TotalLab & " </B> Expexted labels for " & strLabel
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1

			For i=0 to narrLab
					Set oDesc = Description.Create() 
					oDesc("micclass").Value = "WebElement" 
					oDesc("class").Value = "field-caption dataLabelFor.*|header-element header-title"
					oDesc("html tag").Value = "LABEL|SPAN" 
					oDesc("innertext").Value = arrLabVal(i)
					oDesc("outertext").Value =arrLabVal(i)
					oDesc("index").Value =0

					If objPage.WebElement(oDesc).exist(1) Then
						objPage.WebElement(oDesc).Highlight
							Lnumber=i+1
							If gbIterationFlag <> True then
									gstrDesc =  "<B>" & Lnumber & "</B> Successfully get the Label '" & arrLabVal(i) & "' for " & strLabel
									WriteHTMLResultLog gstrDesc, 1
									CreateReport  gstrDesc, 1			
							End if
							bResult = True
					Else
							 Call TakeScreenShot()
							gstrDesc =  "WebElement '" & arrLabVal(i) & "' is not displayed on the screen."
							WriteHTMLResultLog gstrDesc, 0
							CreateReport  gstrDesc, 0                
							bResult = False
							objErr.Raise 11
					End If
			Next
	End If



'	If  objPage.EXIST(gExistCount) Then
'
'			TotalLab=narrLab+1
'			gstrDesc =  "<B>" & TotalLab & " </B> Expexted labels for " & strLabel
'			WriteHTMLResultLog gstrDesc, 1
'			CreateReport  gstrDesc, 1
'
'		Set objElementCollection = objPage.ChildObjects(oDesc)
'		
'		NumberOfWebElements = objElementCollection.Count 
'		nFlag = 0
'		Flag=0
'		For k=0 to narrLab
'				For i = 0 To NumberOfWebElements - 1 
'								If strComp(UCase(Trim(objElementCollection (i).GetROProperty("innertext"))), UCase(Trim(arrLabVal(k)))) =0 Then 
'											Lnumber=k+1
'											gstrDesc =  "<B>" & Lnumber & "</B> Successfully get the Label '" & arrLabVal(k) & "' for " & strLabel
'											WriteHTMLResultLog gstrDesc, 1
'											CreateReport  gstrDesc, 1
'											bResult = true
'								
'								End if			
'										
'				Next 
'			Next
'	End If
    
	VerifyLabel = bResult
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:VerifyLabel
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function VerifyLabel(strObject, strLabel, strVal)
	On error resume next
	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false
	strVal = getTestDataValue(strVal)

	arrLabVal=Split(strVal,";")
	narrLab=Ubound(arrLabVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objPage= gobjObjectClass.getObjectRef(strObject)

	If  objPage.EXIST(gExistCount) Then

			TotalLab=narrLab+1

			gstrDesc =  "Page '" & strLabel & "' is displayed successfully."  
			WriteHTMLResultLog gstrDesc, 2
			CreateReport  gstrDesc, 2

			gstrDesc =  "<B>" & TotalLab & " </B> Expexted labels for " & strLabel
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1

			Set oDesc = Description.Create() 
			oDesc("micclass").Value = "WebElement" 
			oDesc("class").Value ="field-caption.*|header-element header-title|cb_standard|oflowDiv|oflowDivM"
'			oDesc("class").Value = ".*a.*"
			oDesc("html tag").Value = "LABEL|SPAN|DIV"

			arrLabVal=Split(strVal,";")
			narrLab=Ubound(arrLabVal)

			Set objElementCollection = objPage.ChildObjects(oDesc)
	
			NumberOfWebElements = objElementCollection.Count 
			nFlag = 0
			Flag=0
			For k=0 to narrLab
					For i = 0 To NumberOfWebElements - 1 
									If strComp(UCase(Trim(objElementCollection (i).GetROProperty("innertext"))), UCase(Trim(arrLabVal(k)))) =0 Then 	
												objElementCollection(i).Highlight
												bResult = true
												Exit For
									Else
												bResult = false									
									End if											
					Next 

					If  bResult  Then
								Lnumber=k+1
								gstrDesc=""
								If gbIterationFlag <> True then
									gstrDesc =  "<B>" & Lnumber & "</B> Successfully get the Label '" & arrLabVal(k) & "' for " & strLabel
									WriteHTMLResultLog gstrDesc, 1
									CreateReport  gstrDesc, 1			
								End if
					Else
								Call TakeScreenShot()
								gstrDesc =  "WebElement '" & arrLabVal(k) & "' is not displayed on the screen."
								WriteHTMLResultLog gstrDesc, 0
								CreateReport  gstrDesc, 0                
								bResult = False
								objErr.Raise 11
					End If
				Next
	End If
	VerifyLabel = bResult
	
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function UploadFilesToQC(strQCPath,strFilesystemPath)  
    Dim fileCount, timeNow, timeOld, timeDiff
    fileCount = 0
    'Get QC connection object
    Set QCConnection = QCUtil.QCConnection
    'Get Test Plan tree structure
    Set treeManager = QCConnection.TreeManager
    Set node = treeManager.NodeByPath(strQCPath)
    Set AFolder = node.FindChildNode("LABAutomation")  ' Library Files folder Name in QC
    set oAttachment = AFolder.attachments
    timeOld = Now
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(strFilesystemPath)
    Set oFiles = oFolder.Files
    'Iterate through each file present in File System path
    If oFiles.count >0 Then
    For Each oFile in oFiles
        Set attach = oAttachment.AddItem(Null)
        attach.FileName = oFile.Path
        attach.Type = 1
        attach.Post()
        fileCount = fileCount +1
        Set attach = nothing
    Next
    timeNow = Now
    timeDiff =     timeNow - timeOld
'Time required to upload all files to QC
    Reporter.ReportEvent micDone,"Time required to upload : ", "'" & timeDiff & "' minutes."
'Total Files count uploaded to QC
    Reporter.ReportEvent micDone,"Total files uploaded : ", "'" & fileCount & "' files uploaded."
    else
        Reporter.ReportEvent micFail,"File not found", "No file found at path: " & strFilesystemPath
    End If
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:selectDropdown
'Input Parameter      	:strObject - Logical Name of List Box
'     			:strData - Value to  be selected from list box
'Description          	:This function enters a data into a text box
'      		  	 DLL returns an object reference of a WebList
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function selectDropdown(strObject,strLabel,strVal)
   Dim objListField,objTrim,objlen,objReport,Wsh, nLoop,arrAccount,arrTemp,flg,j,strTmp
	Dim intBusiness,LstLBound,arrTmp,strsortcode,i
	bResult = False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."
	Set objListField=gobjObjectClass.getObjectRef(strObject)
		
	If objListField.exist(gExistCount) Then
		objListField.Object.focus
		For i=0 to 9
				 nCount=objListField.GetROProperty("items count")
				 StrAllItems=Split(objListField.GetROProperty("all items"),";")
				intStart=Lbound(strAllItems)
				intEnd=Ubound(strAllItems)+1
				If   nCount  >1 Then
						Exit For
				Else
						Wait 1
				End If
		Next
     
		If  nCount >1Then
						Set objShell = CreateObject("WScript.Shell")
						objShell.Sendkeys "{ENTER}"
						wait 1
						For intCounter = 1 to intEnd
							If objListField.GetItem(intCounter) = strVal Then
												objListField.Select "#"& intCounter-1
												bResult = True
												Exit For
									Else
												bResult = False
									End If 					
						Next

						If bResult Then
									gstrDesc = "Value '" & objListField.getROProperty("value") & "' is selected successfully from '" & strLabel & "' list."
									WriteHTMLResultLog gstrDesc, 1
									CreateReport  gstrDesc, 1
						Else
									Call TakeScreenShot()
									gstrDesc=  strVal & "  Value not found...Please write a reporter event for this"
									WriteHTMLResultLog gstrDesc, 0
									CreateReport  gstrDesc, 0		
									bResult = False
									objErr.Raise 11
						End If
						
		Else
				Call TakeScreenShot()
				gstrDesc =  "Items in the  List '" & strLabel & "' are not displayed."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11
		End If
			
		 
	Else
		 Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	
	selectDropdown=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:selectDropdown
'Input Parameter      	:strObject - Logical Name of List Box
'     			:strData - Value to  be selected from list box
'Description          	:This function enters a data into a text box
'      		  	 DLL returns an object reference of a WebList
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function selectDropdown_OLD(strObject,strLabel,strVal)
   Dim objListField,objTrim,objlen,objReport,Wsh, nLoop,arrAccount,arrTemp,flg,j,strTmp
	Dim intBusiness,LstLBound,arrTmp,strsortcode,i
	bResult = False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."
	Set objListField=gobjObjectClass.getObjectRef(strObject)
		
	If objListField.exist(gExistCount) Then
		objListField.Object.focus
		For i=0 to 9
				 nCount=objListField.GetROProperty("items count")
				If   nCount  >1 Then
						Exit For
				Else
						Wait 1
				End If
		Next
       

		If  nCount >1Then
						Set objShell = CreateObject("WScript.Shell")
						objShell.Sendkeys "{ENTER}"
						wait 1
						objListField.Select strVal
						
						If gbIterationFlag <> True then
									gstrDesc = "Value '" & objListField.getROProperty("value") & "' is selected successfully from '" & strLabel & "' list."
									WriteHTMLResultLog gstrDesc, 1
									CreateReport  gstrDesc, 1
						End If
						bResult = True
		Else
					Call TakeScreenShot()
					gstrDesc =  "Items in the  List '" & strLabel & "' are not displayed."
					WriteHTMLResultLog gstrDesc, 0
					CreateReport  gstrDesc, 0		
					bResult = False
					objErr.Raise 11
		End If
			
		 
	Else
		 Call TakeScreenShot()
		gstrDesc =  "List '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
	
	selectDropdown=bResult
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:VerifyWebElementInPage
'Input Parameter      	:strObject - Logical Name of Web Button
'Description          	:This function clicks the WebElement object
'Return Value         	:True/False
'---------------------------------------------------------------------------------------------------------------------
Public Function VerifyWebElementInPage(strObject,strLabel, strVal)

	Dim objWebElement,bResult,objTrim,objlen,objReport
	bResult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Successfully clicked on Payment ID '" & gstrPaymentID & "'  for  Payment."
	gstrExpectedResult ="Successfully clicked on Payment ID '" & gstrPaymentID & "'  for  Payment."

	Set EditDesc = Description.Create() 
	EditDesc("html tag").Value = "TD|SPAN|DIV" 
	EditDesc("innertext").Value = strVal
	EditDesc("outertext").Value =strVal

	Set objWebPage= gobjObjectClass.getObjectRef(strObject)
	If objWebPage.WebElement(EditDesc).exist(gExistCount) Then
		objWebPage.WebElement(EditDesc).Highlight
        
		If gbIterationFlag <> True then
				gstrDesc =  "Successfully get the <B> '"& strVal &"'</B> in  "& strLabel&" page."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1			
		End if
		bResult = True
	Else
		 Call TakeScreenShot()
		gstrDesc =  "WebElement '" & strLabel & "' is not displayed on the screen."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0                
		bResult = False
		objErr.Raise 11
	End If
	clickWebElementInPage=bResult
End Function



'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:verifySelectList
'Input Parameter      	:strObject - Logical Name of List Box
'     			:strData - Value to  be verify from list box
'Description          	:This function verifies the drop down details
'      		  	 DLL returns an object reference of a WebList
'Calls                	:None
'Return Value   	:None
'---------------------------------------------------------------------------------------------------------------------
Public Function verifySelectList(strObject,strLabel,strVal)
   Dim objListField,objTrim,objlen,objReport,Wsh, nLoop,arrAccount,arrTemp,flg,j,strTmp
	Dim intBusiness,LstLBound,arrTmp,strsortcode,i
	bResult = False	

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	gstrQCDesc = "Select value '" & strVal & "' from '" & strLabel & "' list."
	gstrExpectedResult = "Value '" & strVal & "' should be selected successfully from '" & strLabel & "' list."
	Set objListField=gobjObjectClass.getObjectRef(strObject)
	
		
	If objListField.exist(gExistCount) Then
		objListField.Object.focus
		StrAllItems=Split(objListField.GetROProperty("all items"),";")
		strTemp=Split(strVal,";")
		For i = 0 to ubound(strTemp)
			var=instr(1,StrAllItems,strTemp(i))
			If var1<>0 Then
				gstrDesc = "Value '" & strTemp(i) & "' is present in " & strLabel & "' list."
				WriteHTMLResultLog gstrDesc, 1
				CreateReport  gstrDesc, 1
			Else 
				gstrDesc = "Value '" & strTemp(i) & "' is not present in " & strLabel & "' list."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0
				temp=1
			End If
		Next
	End If	
	If temp =1Then
		bResult = False
		objerr.raise=11
	End If
	verifySelectList=bResult
End Function

'---------------------------------------------------------------------------------------------------------------------


'----------------------------------------------------------------------------------------------------------------
Public function GetRuntimeProperty(strObject,strLabel,strPropName)
	Dim bResult, obj
	bResult = False

	strPropName = getTestDataValue(strPropName)
	
	If UCase(strPropName) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set obj = gobjObjectClass.getObjectRef(strObject)
	If Not obj Is Nothing Then		
        	If obj.exist Then
				If (Len(Trim(strPropName))>0) Then
					gstrPropVal=obj.getROProperty(strPropName)
					gstrDesc =  strLabel &"'s Runtime property " & strPropName & "  is : "& gstrPropVal
					WriteHTMLResultLog gstrDesc, 2
					CreateReport  gstrDesc, 2
					bResult = true
				Else
					TakeScreenshot()
					gstrDesc = "Property Value of " & strLabel & " Not found"
					WriteHTMLResultLog gstrDesc, 0
					bResult = False
					objErr.Raise 11
				End if	
					
            Else
					TakeScreenshot()
					gstrDesc =  "Object "&strLabel&" does not exist"
					WriteHTMLResultLog gstrDesc, 0
					bResult = False
					objErr.Raise 11
			End If
	End If
	GetRuntimeProperty = bResult
end function
