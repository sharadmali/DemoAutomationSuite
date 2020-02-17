'====================================================================================================
'Library File	Name	:Application Utility
'Author			:Sharad Mali
'Creation Date		:
'Description		:Contains definitions of Application specific functions
'====================================================================================================
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Function Name    	: GetCaseIDNumber
'Input Parameter    :strval
'Description        :This Function Gets Application Number
'Calls              :None
'Return Value       :None
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetCaseID(strObjName,strLabel,strVal)
	 Dim arrData1,bResult,strDesc,regEx,Matches,objElm
	strVal = getTestDataValue(strVal)
	 If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	 End If
	 On Error Resume Next
	 bResult=false
    
	 Set objElm = gobjObjectClass.getObjectRef(strObjName)
	
	 If objElm.exist(gExistCount) Then
		   strInnerText=objElm.getROProperty("innertext")     
		  Set regEx = New RegExp   ' Create a regular expression.
		   regEx.Pattern = "\d+"  ' Set pattern.
		   regEx.IgnoreCase = True   ' Set case insensitivity.
		   regEx.Global = True   ' Set global applicability.
           Set Matches = regEx.Execute(strInnerText)
		   if Matches.count>0 Then   ' Execute search.
				gstrCaseID=Matches(0).Value
				If instr(strInnerText,"MOB-MM") > 0 Then
						gstrCaseID="MOB-MM-" & gstrCaseID
				ElseIf instr(strInnerText,"MOB-RBB") > 0 Then
						gstrCaseID="MOB-RBB-" & gstrCaseID
				Else
						gstrCaseID="MOB-" & gstrCaseID
				End If
				

		   End If
		 
		   If gbReporting=true Then
				   If Not isEmpty(gstrCaseID) Then
							 gstrDesc = "Successfully get Case ID :  <B>" & gstrCaseID & "</B>"
							 WriteHTMLResultLog gstrDesc, 1
							CreateReport  gstrDesc, 1
							bResult = True
					else
						  gstrDesc = "Failed to Get ResponseCode Number.............................."
						   WriteHTMLResultLog gstrDesc, 0
						   CreateReport  gstrDesc, 0
						  bResult = False
						
				   End If
		   End If
		Else
			Call TakeScreenShot()
			gstrDesc = "Failed to Get Case ID"
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0
			bResult = False
			objErr.Raise 11
		End IF
	   GetCaseID=bResult
 End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name        	:EnterPegaFormatDate
'Input Parameter    	:strObject - Logical Name of  Webtable
'Description          	:Clicks on specific element from web table
'Calls                	:None
'Return Value         	:True Or False
'---------------------------------------------------------------------------------------------------------------------

Public Function EnterPegaFormatDate(strObject, strLabel, strVal)

	Dim NumberOfWebElements, oDesc, bResult, i, objElementCollection, objTable, nFlag
	bResult = false

	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
		
		MyDate = Date 
		today = FormatDateTime(Date,2)
		chgDate = DateAdd ("d",2,today)
		dayName = WeekDay(chgDate)
		If dayName = 1 Or dayName = 7Then
				chgDate = DateAdd("d",2,chgDate)
		End If
		Futuredate = FormatDateTime(chgDate,2)
		arrDate=Split(Futuredate,"/")
		nDay=arrDate(0)
		nMonth=arrDate(1)
		nYear=arrDate(2)

	Set objEdit = gobjObjectClass.getObjectRef(strObject)
	If  objEdit.EXIST(gExistCount) Then

		If  strVal="DATECWT" Then
					If instr(gstrApplicationURL,"cbtestapp") >0 Then
					strDate= nDay & "/" & nMonth & "/"& nYear &" 11:22"
					ElseIF instr(gstrApplicationURL,"cbdevapp") >0 Then
								strDate= nDay & "/" & nMonth & "/"& nYear &" 11:22"
					ElseIF instr(gstrApplicationURL,"coboo-demo") >0 Then
								strDate = nMonth &"/" & nDay & "/"& nYear &" 11:22 AM"
					Else
								strDate = nMonth &"/" & nDay & "/"& nYear &" 11:22 AM"
					End If
		ElseIf strVal="ROTargetDate" Then
					nDay=nDay+5
					If nDay >30 Then
							nDay=2
					End If
					If gstrUserNamePEGA = "MMBURO" Then
							
'							strDate = nMonth &"/" & nDay & "/"& nYear
							nMonth=nMonth+1
							strDate= nDay & "/" & nMonth & "/"& nYear
					Else
							nMonth=nMonth+1
							strDate= nDay & "/" & nMonth & "/"& nYear
					End If

		ElseIf strVal="ROClientMeetingDate" Then
					If gstrUserNamePEGA = "MMBURO" Then
'							strDate = nMonth &"/" & nDay & "/"& nYear &" 11:22 AM"
							strDate= nDay & "/" & nMonth & "/"& nYear &" 11:22"
					Else
							strDate= nDay & "/" & nMonth & "/"& nYear &" 11:22"
					End If

		ElseIf strVal="ROClarificationDate" Then
					nDay=nDay+2
					If nDay >30 Then
							nDay=1
					End If

					If gstrUserNamePEGA = "MMBURO" Then
'							strDate = nMonth &"/" & nDay & "/"& nYear &" 11:22 AM"
	
							strDate= nDay & "/" & nMonth & "/"& nYear &" 11:22"
					Else

							strDate= nDay & "/" & nMonth & "/"& nYear &" 11:22"
					End If
		ElseIf strVal="ROInterviewDate" Then
					If gstrUserNamePEGA = "MMBURO" Then
'							strDate = nMonth &"/" & nDay & "/"& nYear
							strDate= nDay & "/" & nMonth & "/"& nYear
					Else
							strDate= nDay & "/" & nMonth & "/"& nYear
					End If
		End If


		objEdit.Set strDate
        gstrDesc =  "Successfully entered date for " & strLabel & "is " & strDate 
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = true
	Else
		gstrDesc =  "Failed to enter teh date for " & strLabel
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11

	End If
        
	EnterPegaFormatDate = bResult
	
End Function

'========================================================================================='
'Action For  Select Products
'=========================================================================================''
Public Function SelectProducts(strObject,strLabel,strVal)
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
		Set Dedit=description.Create
		Dedit("micclass").value="WebCheckBox"
		Dedit("html id").value="pyWorkPageProductGroups" & strVal &"ProductSelectedFlag_rdi_" & strVal
		objChkBox.WebCheckBox(Dedit).Set "ON"

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

'========================================================================================='
'Action For  VerifyProductsSelecttion
'========================================================================================='
Public Function VerifyProductsSelecttion(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objPage = gobjObjectClass.getObjectRef(strObject)		
	gstrQCDesc = "Verify Products Selecttion"
	gstrExpectedResult = "Verified Products Selecttion"

	If objPage.exist(gExistCount) Then
		Set Dedit=description.Create
		Dedit("micclass").value="WebCheckBox"
		Dedit("html id").value=".*ProductSelectedFlag_rdi_.*"

		Set Lists = objPage.ChildObjects(Dedit)
        NumberOfLists = Lists.Count()
		For i = 0 To NumberOfLists - 1
					Lists(i).Set "ON"
					Wait 3
					Set Dlst=description.Create
					Dlst("micclass").value="WebList"
					Dlst("html id").value="SubProductSelection"

					Set WebLists = objPage.ChildObjects(Dlst)
					NLists = WebLists.Count()
					NLists=NLists-1
					
					intWebListCount=WebLists(NLists).GetROProperty("items count")
					strAllItems = WebLists(NLists).GetROProperty("all items")	
					

					intWebListCount=intWebListCount-1
					gstrDesc = strLabel & " : Items ->  <B>" & intWebListCount & "</B>.  They are  '" & strAllItems & "' are present in the list."
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
					bResult = True	
					
			Next
	Else
			Call TakeScreenShot()
			gstrDesc = "Items not present in the "& strLabel &" list."
			WriteHTMLResultLog gstrDesc, 0
			CreateReport  gstrDesc, 0		
			bResult = False
			objErr.Raise 11
	End If
	VerifyProductsSelecttion = bresult
End Function
'========================================================================================='
'Action For  Pega Footer
'========================================================================================='
Public Function PegaFooter(strVal)
		strVal = getTestDataValue(strVal)
		
		If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
			Exit Function
		End If
    
		Call Entertext("edtNotes","Notes","Automation Testing Note")
		Call Entertext("edtEffortTime","EffortTime","1")
		Call Clickbutton("btnSubmit","Submit","Click")
'		Call Clickbutton("btnClose","Close","Click")
End Function

'========================================================================================='
'Action For  Verify Items
'========================================================================================='
Public Function VerifyItems(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objListField = gobjObjectClass.getObjectRef(strObject)		
    
	If objListField.exist(1) Then
			intWebListCount=objListField.GetROProperty("items count")
			strAllItems = objListField.GetROProperty("all items")	

			gstrDesc = "Items:  <B>" & intWebListCount & "</B>.  They are  '" & strAllItems & "' are present in the "& strLabel &" list."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = True
	Else
				Call TakeScreenShot()
				gstrDesc = "Items not present in the "& strLabel &" list."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11
	End If
	VerifyItems = bresult
End Function
'========================================================================================='
'Action For  CaseID
'========================================================================================='
Public Function CaseID(strVal)
		Dim objChkBox, bresult
		bresult=False
		strVal = getTestDataValue(strVal)
		If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
				Exit Function
		End If

		If strVal="UPDATE" Then
				Set objConn = CreateObject("ADODB.Connection")
				Set rs = CreateObject("ADODB.Recordset")
				objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb;User Id=;Password=;"
				gstrName="ENV"
				strQuery = "UPDATE TransactionCounter SET CurrentCaseID= '" & gstrCaseID & "' Where ENV = 'PEGA'"
				objConn.BeginTrans
				objConn.Execute(strQuery)
				objConn.CommitTrans	
        		objConn.close
		End If

		If strVal="SELECT" Then
				Set objConn = CreateObject("ADODB.Connection")
				Set rs = CreateObject("ADODB.Recordset")
				objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb;User Id=;Password=;"
				strQuery="select CurrentCaseID from TransactionCounter where ENV='PEGA'"
				rs.open strQuery,objConn,1,1
				If rs.RecordCount>0 Then
						gstrCaseID=rs.Fields("CurrentCaseID")
						rs.Close
						objConn.close
				End If
			End If
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:VerifyCaseInReport
'---------------------------------------------------------------------------------------------------------------------
Function VerifyCaseInReport(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	strArrDetails=Split(strObject,":")

	Set objTable = gobjObjectClass.getObjectRef(strArrDetails(0))		
	Set objLink = gobjObjectClass.getObjectRef(strArrDetails(1))	
	Set objTableHeader=Browser("brwReportsTable").Page("pgReportsTable").WebTable("wbtCurrentStage")


				If  objTable.exist(1) Then
										If  objLink.exist(1) Then
													If IsObject(objLink) Then
																While objTable.GetRowWithCellText(strVal) < 0 And objLink.Exist(1)
																				objLink.Click
																				Wait(5)
																				arrTemp = Split(strObject,":")
																					Set objTable = gobjObjectClass.getObjectRef(strArrDetails(0))		
																					Set objLink = gobjObjectClass.getObjectRef(strArrDetails(1))	
																Wend
													End If
										End If
				
										If objTable.GetRowWithCellText(strVal) > 0 Then
																Set oDesc = Description.Create() 
																oDesc("micclass").Value = "WebElement" 
																Set objElementCollection = objTable.ChildObjects(oDesc)
																NumberOfWebElements = objElementCollection.Count
																For i = 0 To NumberOfWebElements - 1 
																	If Trim(objElementCollection (i).GetROProperty("innertext"))= Trim(strVal) Then 
																							objElementCollection(i).Highlight
																										nRows=objTable.GetRowWithCellText(strVal)
																										nCols=objTable.ColumnCount(1)
																										For k=1 to nCols

																												gstrDesc = objTableHeader.GetCellData(1,k) & "  --> " &  objTable.GetCellData( nRows,k)
																												WriteHTMLResultLog gstrDesc, 1
																												CreateReport  gstrDesc, 1
																										Next	
																												gstrDesc = ""
																												WriteHTMLResultLog gstrDesc, 5
																												CreateReport  gstrDesc, 1

																							bResult = True
'																							Exit for
																	End If
																Next 
															
													
										Else
																bResult = False
												
										End If
				Else
						bResult = False
				End If

			   If bResult Then
							gstrDesc = "Case ID :  <B>" & strVal & "</B>. are present in the "& strLabel &" Report."
							WriteHTMLResultLog gstrDesc, 1
							CreateReport  gstrDesc, 1
							Browser("brwReportsTable").Close
							bResult = True
			   Else
							Call TakeScreenShot()
							gstrDesc = "Case ID :  <B>" & strVal & "</B>. are not  present in the "& strLabel &" Report."
							WriteHTMLResultLog gstrDesc, 0
							CreateReport  gstrDesc, 0		
							bResult = False
							Browser("brwReportsTable").Close
							objErr.Raise 11
			   End If
			
				wait 2
VerifyCaseInReport=bResult
End function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:VerifyAMLcaseSummary
'---------------------------------------------------------------------------------------------------------------------
Function VerifyAMLcaseSummary(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objTable = gobjObjectClass.getObjectRef(strObject)		

	If objTable.exist(1) Then

			nRows=objTable.RowCount
			nCols=objTable.ColumnCount(1)
			For k=2 to nRows
				Browser("brwReportsTable").Page("pgReportsTable").Link("lnkExpandAllGroupHeadings").Click	
				If  Instr(objTable.GetCellData( k,3),"ERROR" )Then
						strData=""
				Else
						strData=objTable.GetCellData( k,3)	
				End If

					gstrDesc = "<B>" & objTable.GetCellData(k,1) & "</B>  --> " &  objTable.GetCellData( k,2) &  "  --> " &  strData
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
				

                    Set WebElm = Browser("brwReportsTable").Page("pgReportsTable").WebTable("wtbAMLCasesSummary").ChildItem(k, 2, "WebElement", 0)
						If  IsObject(WebElm) Then
									strText=WebElm.GetROProperty("innertext")
									If instr(strText,"CDD") OR  instr(strText,"Repair") OR instr(strText,"Pending") OR instr(strText,"Pass") OR instr(strText,"Fail") Then
												WebElm.Highlight
												WebElm.Click
				
											Set objTableHeader=Browser("brwReportsTable").Page("pgReportsTable").WebTable("wbtCurrentStage")
											Set objTableDisplay=Browser("brwReportsTable").Page("pgReportsTable").WebTable("wtbDisplayingRecords")
											
											nRowsD=objTableDisplay.RowCount
											nColsH=objTableHeader.ColumnCount(1)
				
													For i=1 to nRowsD							
																For col=1 to nColsH
																				gstrDesc = "------------------------------>" & objTableHeader.GetCellData(1,col) & "  --> " &  objTableDisplay.GetCellData( i,col)
																				WriteHTMLResultLog gstrDesc, 5
																				CreateReport  gstrDesc, 1
																Next	
													Next
									End If
						End If

						Set WshShell = CreateObject("WScript.Shell")
						wait 2
						WshShell.SendKeys "{BACKSPACE}"
						wait 3
						Set WshShell = Nothing
					


			Next	
			
     Else
				Call TakeScreenShot()
				gstrDesc = "Web Table does not exist"
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11
		End If
		Browser("brwReportsTable").Close

VerifyAMLcaseSummary=bResult
End function

'========================================================================================='
'Action For  EnterCompanyName
'========================================================================================='
Public Function EnterCompanyName(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objEdtField = gobjObjectClass.getObjectRef(strObject)		
    
	If objEdtField.exist(1) Then

		gstrExpectedResult = strVal & " should get typed  in '" & strLabel & "' textbox."
		If Trim(objEdtField.getROProperty("value")) <> ""  Then
			objTextField.set ""
			wait 1
		End If
		objEdtField.click
		Set Wsh = CreateObject("Wscript.Shell")
		wait 2 
		Wsh.SendKeys strVal
		objEdtField.click
		Wsh.SendKeys " "

			Set dobj=Description.Create
			dobj("innertext").value=UCASE(strVal)
			dobj("html tag").value="SPAN"
			dobj("class").value="match-highlight"
			Browser("brwPegaROPortal").Page("pgPegaROPortal").Webelement(dobj).Click	

			gstrDesc = "Successfully entered '" & strVal & "' in '" & strLabel & "' textbox."
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
	EnterCompanyName = bresult
End Function

'========================================================================================='
'Action For  VerifyAdverseMedia
'========================================================================================='
Public Function VerifyAdverseMedia(strObject,strLabel, strVal)
	
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
	objExist.waitProperty "disabled", "0", gExistCount*1000
	If objExist.exist(gExistCount) Then

				Set oDesc1 = Description.Create()
				oDesc1("micclass").Value = "WebElement"
				oDesc1("class").Value = "content layout-content-mimic_a_sentence content-mimic_a_sentence  "
				oDesc1("html tag").Value = "DIV"
				oDesc1("innertext").Value = ".*adverse media.*"
			
			Set ElmCollection =objExist.ChildObjects(oDesc1)
			
			NumberOfElm = ElmCollection.Count
			For i = 0 To NumberOfElm - 1
							strVal= ElmCollection(i).GetROProperty("innertext")
							strArr=Split(strVal,"for")
							strIndividualName=strArr(1)
							strAdverseMedia=strArr(0)

							gstrDesc = "Successfully get the value for " & strLabel & "->" & "Individual Name : <B>" & strIndividualName & "</B> = <B>" & strAdverseMedia &"</B>"
							WriteHTMLResultLog gstrDesc, 4
							CreateReport  gstrDesc, 1
							bResult = True
			Next
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to get the value for " & strLabel
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		objErr.Raise 11
	End If
	VerifyAdverseMedia=bResult
End Function

'========================================================================================='
'Action For  Verify Items
'========================================================================================='
Public Function VerifyItems(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objListField = gobjObjectClass.getObjectRef(strObject)		
    
	If objListField.exist(1) Then
		
		nCount=objListField.GetROProperty("items count")
		If  nCount >1Then
					intWebListCount=objListField.GetROProperty("items count")
					intWebListCount=intWebListCount-1
					strAllItems = objListField.GetROProperty("all items")	
					gstrDesc = strLabel & " : Items ->  <B>" & intWebListCount & "</B>."'  They are  '" & strAllItems & "' are present in the list."
					WriteHTMLResultLog gstrDesc, 1
					CreateReport  gstrDesc, 1
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
				gstrDesc = "Items not present in the "& strLabel &" list."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11
	End If
	VerifyItems = bresult
End Function

'========================================================================================='
'Action For  ClickCompanyName
'========================================================================================='
Public Function ClickCompanyName(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	wait 5
	Set objCmpName = gobjObjectClass.getObjectRef(strObject)		
    
	If objCmpName.exist(1)Then

'			For i=0 to 1
'					Browser("brwPegaROPortal").Refresh
'					Browser("brwPegaROPortal").Sync
'					Set objCmpName = gobjObjectClass.getObjectRef(strObject)	
'					wait 2
'			Next

			objCmpName.Click
			strCMPName=objCmpName.GetROProperty("text")
			gstrDesc =  "Successfully clicked on <B>'" & strCMPName & "'</B> hyperlink."
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
			bResult = True
	Else
				Call TakeScreenShot()
				gstrDesc =  "WebLink '" &strLabel & "' is not displayed on the screen."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11
	End If
	ClickCompanyName = bresult
End Function

'========================================================================================='
'Action For  VerifyPrimaryContact
'========================================================================================='
Public Function VerifyPrimaryContact(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set oDesc1 = Description.Create()
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "content layout-content-inline_grid_33_67 content-inline_grid_33_67"
	oDesc1("html tag").Value = "DIV"

	Set objPage = gobjObjectClass.getObjectRef(strObject)		
    
	If objPage.exist(gExistCount) Then
			Set ElmCollection =objPage.ChildObjects(oDesc1)
			NumberOfElm = ElmCollection.Count
			For i = 0 To NumberOfElm - 1
							strVal= ElmCollection(i).GetROProperty("innertext")
							gstrDesc = "Get the value for " & strLabel & "-> <B>" & strVal & "</B>"
							WriteHTMLResultLog gstrDesc, 4
							CreateReport  gstrDesc, 1
							bResult = True
			Next
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to get the value for " & strLabel
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		objErr.Raise 11
	End If
	VerifyAdverseMedia=bResult
End Function

'-------------------------------------------------------------------------------------------------------------------------------
'Function-Name :ClickOtherActionTable
'Description : This Function verifies table properties
'Output-None
'_________________________________________________________________________
Public Function ClickOtherActionTable(strObject, strLabel, strVal)
	Dim objTable,intRowCnt,intColCnt,bResult
	bResult=False
	wait 1
		
	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objTable = gobjObjectClass.getObjectRef(strObject)
	If  objTable.EXIST(gExistCount) Then
		intRowCnt=objTable.RowCount
		intColCnt=objTable.ColumnCount(intRowCnt)
		intCol=2
		For intRow=1 to intRowCnt
								bResult=True
								If Strcomp(objTable.GetCellData( intRow,intCol) , strVal) = 0 Then
											 gstrDesc =  "Option selected is " & objTable.GetCellData( intRow,intCol)
											 WriteHTMLResultLog gstrDesc, 1
											CreateReport  gstrDesc, 1
											currRow=intRow
											wait 1
											Exit for
									End if 
		Next
	Else 
		bResult=False
		gstrDesc = strLabel & " Table does not exist"
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0
		bResult = False
		objErr.Raise 11
    End If
		
	Set wshobj =Nothing
	ClickOtherActionTable=bResult
End Function

'========================================================================================='
'Action For  GetCaseStatus
'========================================================================================='
Public Function GetCaseStatus(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set oDesc1 = Description.Create()
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "ellipsis"
	oDesc1("html tag").Value = "SPAN"
	Set objPage = gobjObjectClass.getObjectRef(strObject)		
    
	If objPage.exist(gExistCount) Then
			Set ElmCollection =objPage.ChildObjects(oDesc1)
			NumberOfElm = ElmCollection.Count
				Call TakeScreenShot()
			If gbIterationFlag <> True then
					For i = 0 To NumberOfElm - 1
									ElmCollection(i).highlight
									strVal= ElmCollection(i).GetROProperty("innertext")
									gstrDesc = "Get the value for " & strLabel & "-> <B>" & strVal & "</B>"
									WriteHTMLResultLog gstrDesc, 4
									CreateReport  gstrDesc, 1
									bResult = True
					Next
			End If
		
			
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to get the value for " & strLabel
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		objErr.Raise 11
	End If
	GetCaseStatus=bResult
End Function


'========================================================================================='
'Action For  VerifyAverseMediaHits
'========================================================================================='
Public Function VerifyAverseMediaHits(strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	strAM=GetWebelementText("elmAverseMediaHits","Averse Media Hits")

	If strAM=True Then
			Call selectRadioButton("rdCurrentStatusAM","Current Status","Discounted by myself")
			wait 2
			Call Entertext("edtDetailsAdverseMedia","Details AdverseMedia","This is not our Business")
    End If

End Function

'========================================================================================='
'Action For  VerifySanctionHits
'========================================================================================='
Public Function VerifySanctionHits(strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	strAM=GetWebelementText("elmSanctionHits","Averse Media Hits")

	If strAM=True Then
			Call selectRadioButton("rdCurrentStatusS","Current Status","Discounted by me")
			wait 2
			Call Entertext("edtDetailsSanction","DetailsSanction","This is not our Business")
    End If

End Function

'========================================================================================='
'Action For  GetWebelementText
'========================================================================================='
Public Function GetWebelementText(strObject,strLabel)
	Dim objChkBox, bresult
	bresult=False
	Set objElm = gobjObjectClass.getObjectRef(strObject)		
    
	If objElm.exist(1) Then
			strVal= objElm.GetROProperty("innertext")

			If instr(strVal,"Adverse media") > 0  OR instr(strVal,"Sanction hits") > 0  Then
								gstrDesc = "Get the value for " & strLabel & "-> <B>" & strVal & "</B>"
								WriteHTMLResultLog gstrDesc, 4
								CreateReport  gstrDesc, 1
								bR = True
			Else
								gstrDesc = "Get the value for " & strLabel & "-> <B>" & strVal & "</B>"
								WriteHTMLResultLog gstrDesc, 4
								CreateReport  gstrDesc, 1
								bR = False
			End If

			
	Else
			bR = False
	End If

GetWebelementText=bR
End Function


'========================================================================================='
'Action For  RiskChecksComp
'========================================================================================='
Public Function RiskChecksComp(strObject,strLabel, strVal)
	
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
	
	If objExist.exist(gExistCount) Then

				Set oDesc1 = Description.Create()
				oDesc1("micclass").Value = "WebButton"
				oDesc1("name").Value = "Show details"
				oDesc1("visible").Value = True

				
				Set BtnCollection =objExist.ChildObjects(oDesc1)
				NumberOfBtn = BtnCollection.Count


				If NumberOfBtn >0 Then
						
						For i = 0 To NumberOfBtn - 1
									Set oDesc2 = Description.Create()
									oDesc2("micclass").Value = "WebButton"
									oDesc2("name").Value = "Show details"
									oDesc2("visible").Value = True
								
									Set BtnCollection2 =objExist.ChildObjects(oDesc2)
			
									BtnCollection2(i).highlight
									BtnCollection2(i).Click
									k=i+1

									gstrDesc = "<B> Company : "& gstrCompanyName &"</B>"
									WriteHTMLResultLog gstrDesc, 4
									CreateReport  gstrDesc, 1
									wait 5
									strExist=CheckObjectExist("rdCurrentStatusAM","Current Status AdverseMedia","CurrentStatusAdverseMediaExist")
									If strExist Then
											Call selectRadioButton("rdCurrentStatusAM","Current Status","Discounted by myself")
											wait 2
											Call Entertext("edtDetailsAdverseMedia","Details AdverseMedia","This is not our Business")
											Call pressTab()
									End If

									strExist=CheckObjectExist("rdCurrentStatusS","Current Status Sactions","CurrentStatusSactionsExist")
									If strExist Then
											Call selectRadioButton("rdCurrentStatusS","Current Status","Discounted by me")
											wait 2
											Call Entertext("edtDetailsSanction","DetailsSanction","This is not our Business")
											Call pressTab()
									End If
									wait 5

									Call Clicklink("lnkCloseRiskDetails","Close Risk Details","Click")

									wait 5

						Next
						bResult = True

				Else

				End If
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to get the Details for " & strLabel
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		objErr.Raise 11
	End If
	RiskChecksComp=bResult
End Function

'========================================================================================='
'Action For  RiskChecksKAP
'========================================================================================='
Public Function RiskChecksKAP(strObject,strLabel, strVal)
	
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
	
	If objExist.exist(gExistCount) Then

				Set oDesc1 = Description.Create()
				oDesc1("micclass").Value = "WebButton"
				oDesc1("name").Value = "Show details"
				oDesc1("visible").Value = True

				
				Set BtnCollection =objExist.ChildObjects(oDesc1)
				NumberOfBtn = BtnCollection.Count


				If NumberOfBtn >0 Then
						
						For i = 0 To NumberOfBtn - 1
									Set oDesc2 = Description.Create()
									oDesc2("micclass").Value = "WebButton"
									oDesc2("name").Value = "Show details"
									oDesc2("visible").Value = True

								
									Set BtnCollection2 =objExist.ChildObjects(oDesc2)
			
									BtnCollection2(i).highlight
									BtnCollection2(i).Click
									k=i+1

									gstrDesc = "<B> Individual : "& k &"</B>"
									WriteHTMLResultLog gstrDesc, 4
									CreateReport  gstrDesc, 1
                                    
									strExist=CheckObjectExist("rdCurrentStatusP","Current Status PEPs","CurrentStatusPEPsExist")
									If  strExist Then
												wait 2
												Call selectRadioButton("rdCurrentStatusP","Current Status PEPs","Discounted by me")
'												Call selectRadioButton("rdPEPRiskLevel","PEP Risk Level","Medium risk PEP")
												wait 2
												Call Entertext("edtDetailsPEPs","Details PEPs","This is not our KAP")
									End IF
		
									strExist=CheckObjectExist("rdCurrentStatusS","Current Status Sanction","CurrentStatusSanctionExist")
									If  strExist Then
												wait 2
												Call selectRadioButton("rdCurrentStatusS","Current Status Sanction","Discounted by myself")
												wait 2
												Call Entertext("edtDetailsSanction","Details Sanction","This is not our KAP")
									End If
		
									strExist=CheckObjectExist("rdCurrentStatusAM","Current Status Adverse Media","CurrentStatusAdverseMediaExist")
									If  strExist Then
												wait 2
												Call selectRadioButton("rdCurrentStatusAM","Current Status Adverse Media","Discounted by myself")
												wait 2
												Call Entertext("edtDetailsAdverseMedia","Details AdverseMedia","This is not our KAP")
									End If
									wait 5
									Call Clicklink("lnkCloseRiskDetails","Close Risk Details","Click")
									wait 5
						Next
						bResult = True

				Else

				End If
	Else
		Call TakeScreenShot()
		gstrDesc ="Failed to get the Details for " & strLabel
		WriteHTMLResultLog gstrDesc, 4
		CreateReport  gstrDesc, 1    		
		bResult = False
		objErr.Raise 11
	End If
	RiskChecksKAP=bResult
End Function

'========================================================================================='
'Action For  ProductsSelecttion CWT
'========================================================================================='

Public Function ProductsSelectionCWT(strObject,strLabel,strVal)

	On error resume next
   	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objPage = gobjObjectClass.getObjectRef(strObject)		
	If objPage.Exist(gExistCount) Then

	ProductArray = Split(strVal,";")
	SubProductArray = Split(strSubProduct,";")
	FeatureArray = Split(strFeature,";")

	   For count1= 0 to Ubound(ProductArray)
				Set BtnDesc = Description.Create() 
				BtnDesc("Name").Value = "Add account" 
				BtnDesc("html tag").Value = "BUTTON" 
				BtnDesc("Index").Value =count1

'				cDrop=count1+2

				Set DropDesc = Description.Create() 
				DropDesc("class").Value = "expandRowDetails" 
				DropDesc("html tag").Value = "A" 
'				DropDesc("Index").Value =cDrop
				DropDesc("Index").Value =count1

				Set DropSubAccnt = Description.Create() 
				DropSubAccnt("micclass").Value = "WebList" 
				DropSubAccnt("html tag").Value = "SELECT" 
				DropSubAccnt("html id").Value = "SubProductSelection" 
				DropSubAccnt("Index").Value =count1

				Set lstDescAddFeature = Description.Create() 
				lstDescAddFeature("micclass").Value = "WebList" 
				lstDescAddFeature("html tag").Value = "SELECT" 
				lstDescAddFeature("html id").Value = "FeatueSelection" 
				lstDescAddFeature("Index").Value =count1

				Set BtnAddFeatureDesc = Description.Create() 
				BtnAddFeatureDesc("Name").Value = "Add feature" 
				BtnAddFeatureDesc("html tag").Value = "BUTTON" 
				BtnAddFeatureDesc("Index").Value =count1
				
				If Instr(ProductArray(count1),"Current Account")>0Then
								Call setCheckBox("chkCurrentAccount","Current Account","ON")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "#5"
																gstrDesc = "Value 'EBT' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 5
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Call ClickWebelement("elmPCMPPrimaryAccount","Primary Account","Click")
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(lstDescAddFeature).Select "Debit Card"
																gstrDesc = "Value 'Debit Card' is selected successfully from 'Add Feature' list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 1
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnAddFeatureDesc).Click
								wait 1
								
								Call Entertext("edtFacilityToBeProvided","Facility To Be Provided","TEST1")
								Call Entertext("edtFacilityNote","Facility Note","TEST2")
								Call clickButton("btnSaveProduct","Save","Click")
								wait 2
				End If	

				If Instr(ProductArray(count1),"Ancillary Limits")>0Then
								Call setCheckBox("chkAncillaryLimit","Ancillary Limits","ON")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "BACS"
																gstrDesc = "Value 'BACS' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
'								wait 2
'								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
'								wait 2
'								Call EnterText("edtAccountName","Banking Account Name","Account Name")
'								wait 2
'								Call clickButton("btnSaveProduct","Save","Click")
'								wait 2
				End If							

				If Instr(ProductArray(count1),"PartnerProducts")>0Then
							Call setCheckBox("chkAncillaryLimit","Ancillary Limit","ON")
							wait 4
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Cardnet"
															gstrDesc = "Value 'Cardnet' is selected successfully from 'Sub Account list'."
															WriteHTMLResultLog gstrDesc, 1
															CreateReport  gstrDesc, 1
							wait 2
							Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
							wait 2
							Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
							wait 2
							Call EnterText("edtAccountName","Banking Account Name","Account Name")
							wait 2
							Call clickButton("btnSaveProduct","Save","Click")
							wait 2
			End If			

			If Instr(ProductArray(count1),"Client Account")>0Then
								Call setCheckBox("chkClientAccount","Current Account","ON")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Client Call Account"
																gstrDesc = "Value 'Client Call Account' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Call clickButton("btnSaveProduct","Save","Click")
								wait 2
			End If	

			If Instr(ProductArray(count1),"Currency Account")>0Then
								Call setCheckBox("chkCurrencyAccount","Currency Account","ON")
								wait 5
								Browser("brwPegaCaseManagerPortal").Refresh

								wait 10
								Call selectlist("lstSterlingAccount","Sterling Account","CurrentAccount-Account Name")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Euro"
																gstrDesc = "Value 'Euro' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(lstDescAddFeature).Select "SMS"
																gstrDesc = "Value 'SMS' is selected successfully from 'Add Feature' list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 1
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnAddFeatureDesc).Click

								Call Entertext("edtFacilityToBeProvided","Facility To Be Provided","TEST1")
								Call Entertext("edtFacilityNote","Facility Note","TEST2")
								Call clickButton("btnSaveProduct","Save","Click")
								wait 5
								Call Selectlist("lstMaintenanceChargesAccount","Maintenance Charges Account","CurrentAccount-Account Name")
'								Call EnterText("edtSITCode","edtSITCode","1234567")	
				End If
		
				If Instr(ProductArray(count1),"Deposit Account")>0Then
							   Call setCheckBox("chkDepositAccount","Deposit Account","ON")
							   Wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Business Instant Access Account"
																gstrDesc = "Value 'Business Instant Access Account' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
												wait 2
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
												wait 2
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
												wait 2
												Call EnterText("edtAccountName","Banking Account Name","Account Name")
												wait 2
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(lstDescAddFeature).Select "SMS"
																				gstrDesc = "Value 'SMS' is selected successfully from 'Add Feature' list'."
																				WriteHTMLResultLog gstrDesc, 1
																				CreateReport  gstrDesc, 1
												wait 1
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnAddFeatureDesc).Click
				
												Call Entertext("edtFacilityToBeProvided","Facility To Be Provided","TEST1")
												Call Entertext("edtFacilityNote","Facility Note","TEST2")
												Call clickButton("btnSaveProduct","Save","Click")
												wait 2
				End If

				If Instr(ProductArray(count1),"eBanking")>0Then
							   Call setCheckBox("chkeBanking","eBanking","ON")
							   Wait 5
											
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Online for Business"
														gstrDesc = "Value 'Online for Business' is selected successfully from 'Sub Account list'."
														WriteHTMLResultLog gstrDesc, 1
														CreateReport  gstrDesc, 1
											wait 2
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
											wait 2
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
											wait 2
											Call EnterText("edtAccountName","Banking Account Name","Account Name")
											wait 2
											Call clickButton("btnSaveProduct","Save","Click")
											wait 2
				End if 
    Next 
	End If
End Function

'========================================================================================='
'Action For  ProductsSelecttion CWT
'========================================================================================='

Public Function ProductseDeletionCWT(strObject,strLabel,strVal)

   	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objPage = gobjObjectClass.getObjectRef(strObject)		
	If objPage.Exist(gExistCount) Then

	ProductArray = Split(strVal,";")
	SubProductArray = Split(strSubProduct,";")
	FeatureArray = Split(strFeature,";")

	   For count1= 0 to Ubound(ProductArray)
				Set BtnDesc = Description.Create() 
				BtnDesc("Name").Value = "Add account" 
				BtnDesc("html tag").Value = "BUTTON" 
				BtnDesc("Index").Value =count1

'				cDrop=count1+2

				Set DropDesc = Description.Create() 
				DropDesc("class").Value = "expandRowDetails" 
				DropDesc("html tag").Value = "A" 
'				DropDesc("Index").Value =cDrop
				DropDesc("Index").Value =count1

				Set DropSubAccnt = Description.Create() 
				DropSubAccnt("micclass").Value = "WebList" 
				DropSubAccnt("html tag").Value = "SELECT" 
				DropSubAccnt("html id").Value = "SubProductSelection" 
				DropSubAccnt("Index").Value =count1

				Set lstDescAddFeature = Description.Create() 
				lstDescAddFeature("micclass").Value = "WebList" 
				lstDescAddFeature("html tag").Value = "SELECT" 
				lstDescAddFeature("html id").Value = "FeatueSelection" 
				lstDescAddFeature("Index").Value =count1

				Set BtnAddFeatureDesc = Description.Create() 
				BtnAddFeatureDesc("Name").Value = "Add feature" 
				BtnAddFeatureDesc("html tag").Value = "BUTTON" 
				BtnAddFeatureDesc("Index").Value =count1
				
				If Instr(ProductArray(count1),"Current Account")>0Then
						Call setCheckBox("chkCurrentAccount","Current Account","OFF")				
				End If	
				
				If Instr(ProductArray(count1),"Ancillary Limits")>0Then
						Call setCheckBox("chkAncillaryLimit","Ancillary Limits","ONOFF")				
				End If							
				
				If Instr(ProductArray(count1),"PartnerProducts")>0Then
						Call setCheckBox("chkAncillaryLimit","Ancillary Limit","OFF")				
				End If			
				
				If Instr(ProductArray(count1),"Client Account")>0Then
						Call setCheckBox("chkClientAccount","Current Account","OFF")			
				End If	
				
				If Instr(ProductArray(count1),"Currency Account")>0Then
						Call setCheckBox("chkCurrencyAccount","Currency Account","OFF")			
				End If
				
				If Instr(ProductArray(count1),"Deposit Account")>0Then
						Call setCheckBox("chkDepositAccount","Deposit Account","OFF")				
				End If
				
				If Instr(ProductArray(count1),"eBanking")>0Then
						Call setCheckBox("chkeBanking","eBanking","OFF")			
				End if 
    Next 
	End If
End Function

'========================================================================================='
'Action For  ProductsSelecttion RO
'========================================================================================='

Public Function ProductsSelectionRO(strObject,strLabel,strVal)

   	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
    
	Set objPage = gobjObjectClass.getObjectRef(strObject)		
	If objPage.Exist(gExistCount) Then

	ProductArray = Split(strVal,";")
	SubProductArray = Split(strSubProduct,";")
	FeatureArray = Split(strFeature,";")

	   For count1= 0 to Ubound(ProductArray)
				Set BtnDesc = Description.Create() 
				BtnDesc("Name").Value = "Add account" 
				BtnDesc("html tag").Value = "BUTTON" 
				BtnDesc("Index").Value =count1

				Set DropDesc = Description.Create() 
				DropDesc("class").Value = "expandRowDetails" 
				DropDesc("html tag").Value = "A" 
				count1Drop=count1+2
				DropDesc("Index").Value =count1Drop

				Set DropSubAccnt = Description.Create() 
				DropSubAccnt("micclass").Value = "WebList" 
				DropSubAccnt("html tag").Value = "SELECT" 
				DropSubAccnt("html id").Value = "SubProductSelection" 
				DropSubAccnt("Index").Value =count1

				Set lstDescAddFeature = Description.Create() 
				lstDescAddFeature("micclass").Value = "WebList" 
				lstDescAddFeature("html tag").Value = "SELECT" 
				lstDescAddFeature("html id").Value = "FeatueSelection" 
				lstDescAddFeature("Index").Value =count1

				Set BtnAddFeatureDesc = Description.Create() 
				BtnAddFeatureDesc("Name").Value = "Add feature" 
				BtnAddFeatureDesc("html tag").Value = "BUTTON" 
				BtnAddFeatureDesc("Index").Value =count1
				
				 If Instr(ProductArray(count1),"eBanking")>0Then
'							   Call setCheckBox("chkeBankingRO","eBanking","ON")
							   Call ClickSiblingPrevious("eBankingRO","eBanking","Click")
							   Wait 5
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Online for Business"
														gstrDesc = "Value 'Online for Business' is selected successfully from 'Sub Account list'."
														WriteHTMLResultLog gstrDesc, 1
														CreateReport  gstrDesc, 1
											wait 2
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
											wait 4
											Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
											wait 2
											Call EnterText("edtAccountName","Banking Account Name","Account Name")
											wait 2
											Call clickButton("btnSaveProduct","Save","Click")
											wait 2
				End if 

				If Instr(ProductArray(count1),"Deposit Account")>0Then
'							   Call setCheckBox("chkDepositAccountRO","Deposit Account","ON")
							   Call ClickSiblingPrevious("DepositAccountRO","Deposit Account RO","Click")
							   Wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Business Instant Access Account"
																gstrDesc = "Value 'Business Instant Access Account' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
												wait 2
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
												wait 2
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
												wait 2
												Call EnterText("edtAccountName","Banking Account Name","Account Name")
												wait 2
												Call clickButton("btnSaveProduct","Save","Click")
												wait 2
				End if

				If Instr(ProductArray(count1),"Current Account")>0Then
'								Call setCheckBox("chkCurrentAccountRO","Current Account","ON")
								 Call ClickSiblingPrevious("CurrentAccountRO","Current Account RO","Click")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "#5"
																gstrDesc = "Value 'EBT' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Call ClickWebelement("elmPCMPPrimaryAccount","Primary Account","Click")
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(lstDescAddFeature).Select "Debit Card"
																gstrDesc = "Value 'Debit Card' is selected successfully from 'Add Feature' list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 1
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnAddFeatureDesc).Click

								Call Entertext("edtFacilityToBeProvided","Facility To Be Provided","TEST1")
								Call Entertext("edtFacilityNote","Facility Note","TEST2")
								Call clickButton("btnSaveProduct","Save","Click")
								wait 2
				End If					

				If Instr(ProductArray(count1),"Currency Account")>0Then
'								Call setCheckBox("chkCurrencyAccountRO","Current Account","ON")
								Call ClickSiblingPrevious("CurrencyAccountRO","Currency Account RO","Click")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Euro"
																gstrDesc = "Value 'Euro' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Call clickButton("btnSaveProduct","Save","Click")
								wait 2
								Call EnterText("edtSITCode","edtSITCode","1234567")	
				End if
				
				If Instr(ProductArray(count1),"Client Account")>0Then
'								Call setCheckBox("chkClientAccountRO","Current Account","ON")
								Call ClickSiblingPrevious("ClientAccountRO","Client Account RO","Click")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Client Call Account"
																gstrDesc = "Value 'Client Call Account' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Call clickButton("btnSaveProduct","Save","Click")
								wait 2
				End If			

				If Instr(ProductArray(count1),"Ancillary limits")>0Then
'								Call setCheckBox("chkAnsillaryProductsRO","Ansillary Products","ON")
								Call ClickSiblingPrevious("AncillaryLimitsRO","Ancillary Limits","Click")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "BACS"
																gstrDesc = "Value 'BACS' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
'								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
'								wait 2
'								Call EnterText("edtAccountName","Banking Account Name","Account Name")
'								wait 2
'								Call clickButton("btnSaveProduct","Save","Click")
'								wait 2
				End If		

				If Instr(ProductArray(count1),"Partner Products")>0Then
'								Call setCheckBox("chkPartnerProductsRO","Partner Products","ON")
								Call ClickSiblingPrevious("PartnerProductsRO","Partner Products  RO","Click")
								wait 4
												Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebList(DropSubAccnt).Select "Cardnet"
																gstrDesc = "Value 'Cardnet' is selected successfully from 'Sub Account list'."
																WriteHTMLResultLog gstrDesc, 1
																CreateReport  gstrDesc, 1
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebButton(BtnDesc).Click
								wait 2
								Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement(DropDesc).Click
								wait 2
								Call EnterText("edtAccountName","Banking Account Name","Account Name")
								wait 2
								Call clickButton("btnSaveProduct","Save","Click")
								wait 2
				End If				
    Next 
	End If
End Function


'========================================================================================='
'Action For setAllCheckBoxMandateSent
'========================================================================================='
Public Function setAllCheckBoxMandateSent(strObject,strLabel,strVal)
	Dim objChk, bresult,objPage,arrChk,nLoop
	 
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

    Set objChk = Description.Create
	objChk("micclass").Value = "WebCheckBox" 
	objChk("name").Value = ".*MandateSent|.*KAPEmailSent|.*AuthorisedSignatory"
	
	Set objPage = gobjObjectClass.getObjectRef(strObject)
	Set arrChk = objPage.ChildObjects(objChk)
	If objPage.Exist(gExistCount) And arrChk.Count > 0 Then
		For nLoop = 0 to arrChk.count-1
				Set objPage = gobjObjectClass.getObjectRef(strObject)
				Set objChk = Description.Create
				objChk("micclass").Value = "WebCheckBox" 
				objChk("name").Value = ".*MandateSent|.*KAPEmailSent|.*AuthorisedSignatory"
				Set arrChk = objPage.ChildObjects(objChk)

				strName=arrChk(nLoop).GetROProperty("name")
				If UCase(strVal) = "OFF" Then
						arrChk(nLoop).set  "OFF"
				Else
						arrChk(nLoop).set  "ON"
						wait 2

						If  nLoop=0 Then
'								gstrDesc =  "Successfully set value " & strVal & " for  check box  " & strName & " available on '" & strLabel & "' page."		
								gstrDesc =  "Successfully set value <B>" & strVal & " </B>for  check box  <B> Mandate configured and sent </B> available on '" & strLabel & "' page."	
								WriteHTMLResultLog gstrDesc, 1
								CreateReport  gstrDesc, 1
						ElseIf nLoop=1 Then
								gstrDesc =  "Successfully set value <B>" & strVal & " </B>for  check box  <B> KAP Email sent   </B> available on '" & strLabel & "' page."	
								WriteHTMLResultLog gstrDesc, 1
								CreateReport  gstrDesc, 1
						Else
								gstrDesc =  "Successfully set value <B>" & strVal & " </B>for  check box  <B> Authorised signatory   </B> available on '" & strLabel & "' page."	
								WriteHTMLResultLog gstrDesc, 1
								CreateReport  gstrDesc, 1
						End If
				End If		
		Next
		
	Else
		Call TakeScreenShot()
		gstrDesc =  "Page  '" & strLabel & "' is not displayed on screen."		
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		objErr.Raise 11
	End If
    setAllCheckBoxMandateSent = bresult
End Function

'========================================================================================='
'Action For setAllCheckBox for Mandate Received
'========================================================================================='
Public Function setAllCheckBoxMandateAccepted(strObject,strLabel,strVal)
	Dim objChk, bresult,objPage,arrChk,nLoop
	 
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	wait 5
    Set objChk = Description.Create
	objChk("micclass").Value = "WebCheckBox" 
	objChk("name").Value = ".*MandateAccepted"
	
	Set objPage = gobjObjectClass.getObjectRef(strObject)
	Set arrChk = objPage.ChildObjects(objChk)
	If objPage.Exist(gExistCount) And arrChk.Count > 0 Then
		For nLoop = 0 to arrChk.count-1
				Set objPage = gobjObjectClass.getObjectRef(strObject)
				Set objChk = Description.Create
				objChk("micclass").Value = "WebCheckBox" 
				objChk("name").Value = ".*MandateAccepted"
				Set arrChk = objPage.ChildObjects(objChk)

				strName=arrChk(nLoop).GetROProperty("name")
				If UCase(strVal) = "OFF" Then
						arrChk(nLoop).set  "OFF"
				Else
						arrChk(nLoop).set  "ON"
						wait 2
								gstrDesc =  "Successfully set value <B>" & strVal & " </B>for  check box  <B> Mandate accepted    </B> available on '" & strLabel & "' page."	
								WriteHTMLResultLog gstrDesc, 1
								CreateReport  gstrDesc, 1
				End If		
		Next
		
	Else
		Call TakeScreenShot()
		gstrDesc =  "Page  '" & strLabel & "' is not displayed on screen."		
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		objErr.Raise 11
	End If
    setAllCheckBoxMandateReceived = bresult
End Function

'========================================================================================='
'Action For  IndividualDetailsCWT
'========================================================================================='
Public Function IndividualDetailsCWT(strObject,strLabel,strVal)
   On error resume next
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objTable = gobjObjectClass.getObjectRef(strObject)		
    
	If objTable.Exist(1) Then
	
			If  strVal="UPDATE" Then 
						
    						Set oDesc = Description.Create() 
							oDesc("micclass").Value = "WebElement" 
							oDesc("class").Value = "expandRowDetails" 
							oDesc("html tag").Value = "A" 

							Set objElementCollection=objTable.ChildObjects(oDesc)
							NumberOfWebElements = objElementCollection.Count 

								For ntRow = 0 To NumberOfWebElements-1
'									For ntRow = 0 To 0
												wait 3
												Set objTable1 = gobjObjectClass.getObjectRef(strObject)
												Set oDesc1 = Description.Create() 
												oDesc1("micclass").Value = "WebElement" 
												oDesc1("class").Value = "expandRowDetails" 
												oDesc1("html tag").Value = "A" 

												Set objElementCollection1=objTable1.ChildObjects(oDesc)
												NumberOfWebElements1 = objElementCollection1.Count 

												objElementCollection1(ntRow).highlight							
												objElementCollection1(ntRow).FireEvent "Click"
												wait 3

												If Not  ntRow =0 And Not  ntRow =1 Then
																Call setCheckBox("chkIndividualKAPCWT","Individual KAP CWT","OFF")
																Call Clickbutton("btnSaveCWT","Save","Click")
												Else
'																If  ntRow=0 Then
'																		Call EnterText("edtTitleCWT","Title","Mr")
'																Else
'																		Call EnterText("edtTitleCWT","Title","Miss")
'																End If
'																
'																If ntRow=0  Then
'																	Call EnterText("edtFirstNameCWT","First Name","Sharad")
'																Else
'																	Call EnterText("edtFirstNameCWT","First Name","Dipti")
'																End If
'									
'																If ntRow=0  Then
'																		Call EnterText("edtLastNameCWT","Last Name","Mali")
'																Else
'																		Call EnterText("edtLastNameCWT","Last Name","Bhushetty")
'																End If
									
																If ntRow=0  Then
																		Call selectRadioButton("rdGenderCWT","Gender","Male")
																Else
																		Call selectRadioButton("rdGenderCWT","Gender","Female")
																End If
																Call VerifyDisplayProperty("lstNationality","Nationality","value")
																Call VerifyDisplayProperty("edtDateOfBirth","Date Of Birth","value")
																'Call VerifyDisplayProperty("edtOccupation","Occupation","value")
																If gstrCurID=10 OR gstrCurID=11 OR gstrCurID=23 OR gstrCurID=33 OR gstrCurID=35 OR gstrCurID=79 OR gstrCurID=83 OR gstrCurID=88 OR gstrCurID=93 OR gstrCurID=98 OR gstrCurID=603 OR gstrCurID=604 OR gstrCurID=606 Then ' Soletrader
				
																Else
																	Call VerifyDisplayProperty("edtPositionInCompany","Position In Company","value")
																End If									
																Call VerifyItems("lstCountryOfResidence","Country Of Residence","Verify")
				
																If gstrCurID=11 OR gstrCurID=23 OR gstrCurID=606 Then
				
																Else   													
																			Call VerifyItems("lstRole","Role","Verify")
																			Call selectItemInList("lstRole","Role","2")
																End If
				
																If ntRow=0 Then
																				Call EnterText("edtEmailRO","Email","sharad.mali@yopmail.com")
																ElseIf ntRow=1 Then
																				Call EnterText("edtEmailRO","Email","dipti.bhushetty@yopmail.com")
																End If
				
																If ntRow=0 Then
																				Call Entertext("edtMobileNumber","Mobile Number","07438116882")
																ElseIf ntRow=1 Then
																				Call Entertext("edtMobileNumber","Mobile Number","07459838205")
																End If
				
																If ntRow=0 Then
																				Call Entertext("edtPrimaryContactNumber","Primary Contact Number","7438116882")
																ElseIf ntRow=1 Then
																				Call Entertext("edtPrimaryContactNumber","Primary Contact Number","07459838205")
																End If
				
																If ntRow=0 Then
																				Call setCheckBox("chkPrimaryContact","Primary Contact","ON")
																End If
																If  ntRow =0 Or ntRow =1Then
																				Call setCheckBox("chkAuthorisedSignatory","Authorised Signatory","ON")
																				Call setCheckBox("chkValidatedCWT","chkValidated","ON")
																End If

																Call Clickbutton("btnSaveCWT","Save","Click")
											   End If
                               Next

			ElseIf  strVal="ADD" Then 

					For i=0 to 1
							Call ClickLink("lnkAddIndividualCWT","Add Individual Link","Click")		
							wait 2
							If  i=0 Then
									Call EnterText("edtTitleCWT","Title","Mr")
							Else
									Call EnterText("edtTitleCWT","Title","Miss")
							End If
							
							If i=0  Then
								Call EnterText("edtFirstNameCWT","First Name","Sharad")
							Else
								Call EnterText("edtFirstNameCWT","First Name","Dipti")
							End If

							If i=0  Then
									Call EnterText("edtLastNameCWT","Last Name","Mali")
							Else
									Call EnterText("edtLastNameCWT","Last Name","Bhushetty")
							End If

							If i=0  Then
									Call selectRadioButton("rdGenderCWT","Gender","Male")
							Else
									Call selectRadioButton("rdGenderCWT","Gender","Female")
							End If

							If i=0  Then
									Call EnterText("edtEmailCWT","Email","sharad.mali@abc.com")
							Else
									Call EnterText("edtEmailCWT","Email","sharad.mali@abc.com")
							End If

							If i=0  Then
									Call EnterText("edtMobileCWT","Mobile","7438116882")
							Else
									Call EnterText("edtMobileCWT","Mobile","07459838205")
							End If

							If i=0  Then
									Call EnterText("edtPrimaryNumberCWT","Primary Number","7438116882")
							Else
									Call EnterText("edtPrimaryNumberCWT","Primary Number","07459838205")
							End If

							Call SelectList("lstNationality","Nationality","British")
							Call SelectList("lstDualNationality","Dual Nationality","British")
							Call VerifyItems("lstCountryOfResidence","Country Of Residence","Verify")
							Call Entertext("edtOccupation","Occupation","Director")
													
							If gstrCurID=11 OR gstrCurID=23 OR gstrCurID=606 Then
				
							Else   													
										Call VerifyItems("lstRole","Role","Verify")
										Call selectItemInList("lstRole","Role","2")
							End If

							If gstrCurID=10 OR gstrCurID=11 OR gstrCurID=23 OR gstrCurID=33 OR gstrCurID=35 OR gstrCurID=79 OR gstrCurID=83 OR gstrCurID=88 OR gstrCurID=93 OR gstrCurID=98 OR gstrCurID=603 OR gstrCurID=604 OR gstrCurID=606 OR  gstrCurID=158 OR gstrCurID=163 OR gstrCurID=168 OR gstrCurID=173Then ' Soletrader
				
							Else
								Call Entertext("edtPositionInCompany","Position In Company","Director")
							End If	


							Call setCheckBox("chkIndividualKAPCWT","Individual KAP CWT","ON")
							If i=0 Then
								Call setCheckBox("chkPrimaryContact","Primary Contact","ON")
							End If							
							Call setCheckBox("chkValidatedCWT","Validated","ON")
							Call setCheckBox("chkAuthorisedSignatory","Authorised Signatory","ON")
							Call ClickButton("btnSaveCWT","Save Individual","Click")	
					Next
									
			ElseIf strVal="VERIFY" Then 
						
    						Set oDesc = Description.Create() 
							oDesc("micclass").Value = "WebElement" 
							oDesc("class").Value = "expandRowDetails" 
							oDesc("html tag").Value = "A" 

							Set objElementCollection=objTable.ChildObjects(oDesc)
							NumberOfWebElements = objElementCollection.Count 

								For ntRow = 0 To NumberOfWebElements-1
												wait 3
												Set objTable1 = gobjObjectClass.getObjectRef(strObject)
												Set oDesc1 = Description.Create() 
												oDesc1("micclass").Value = "WebElement" 
												oDesc1("class").Value = "expandRowDetails" 
												oDesc1("html tag").Value = "A" 

												Set objElementCollection1=objTable1.ChildObjects(oDesc)
												NumberOfWebElements1 = objElementCollection1.Count 

												objElementCollection1(ntRow).highlight							
												objElementCollection1(ntRow).FireEvent "Click"
												wait 3
												
												Call VerifyDisplayProperty("lstNationality","Nationality","value")
												'	Call VerifySibling("elmYearAndMonthOfBirth","Year And Month Of Birth","YearAnMonthOfBirthVerify")
												Call VerifyDisplayProperty("edtDateOfBirth","Date Of Birth","value")
												Call VerifyDisplayProperty("edtOccupation","Occupation","value")
												Call VerifyDisplayProperty("edtPositionInCompany","Position In Company","value")
												Call VerifyDisplayProperty("lstNationality","Nationality","value")
												Call VerifyDisplayProperty("lstDualNationality","Dual Nationality","value")
												Call VerifyDisplayProperty("lstCountryOfResidence","Country Of Residence","value")
												Call VerifyDisplayProperty("lstRole","Role","value")
												Call VerifyDisplayProperty("edtEMail","EMail","value")
												Call VerifyDisplayProperty("edtMobileNumber","Mobile Number","value")
												Call VerifyDisplayProperty("edtPrimaryContactNumber","Primary Contact Number","value")
												Call VerifyDisplayProperty("chkPrimaryContact","Primary Contact","checked")
												Call VerifyDisplayProperty("chkAuthorisedSignatory","Authorised Signatory","checked")
												Call VerifyDisplayProperty("chkValidatedCWT","Validated","checked")
												Call Clickbutton("btnSaveCWT","Save","Click")
                               Next
				ElseIf strVal="DELETE" Then 
						
    						Set oDesc = Description.Create() 
							oDesc("micclass").Value = "WebElement" 
							oDesc("class").Value = "iconDelete" 
							oDesc("html tag").Value = "A" 

							Set objElementCollection=objTable.ChildObjects(oDesc)
							NumberOfWebElements = objElementCollection.Count 

								For ntRow = 0 To NumberOfWebElements-1
												wait 3
												Set objTable1 = gobjObjectClass.getObjectRef(strObject)
												Set oDesc1 = Description.Create() 
												oDesc1("micclass").Value = "WebElement" 
												oDesc1("class").Value = "iconDelete" 
												oDesc1("html tag").Value = "A" 

												Set objElementCollection1=objTable1.ChildObjects(oDesc)
												NumberOfWebElements1 = objElementCollection1.Count 
												If ntRow=1 Then
														ntRow=0
												End If
												objElementCollection1(ntRow).highlight							
												objElementCollection1(ntRow).FireEvent "Click"
												wait 3
												Call ClickWebElement("elmReasonForDeletionKAP","Reason For Deletion KAP","Click")
												wait 3
												Call clickbutton("btnDeleteIndividual", "Delete Individual","Click")
												wait 5
												ntRow1=ntRow+1
												gstrDesc = "<B>KAP deleted successfully.</B>"
												WriteHTMLResultLog gstrDesc, 1
												CreateReport  gstrDesc, 1

                               Next
				ElseIf  strVal="ADDNEW" Then 

					For i=0 to 62
							Call ClickLink("lnkAddIndividualCWT","Add Individual Link","Click")		
							wait 2
							Call EnterText("edtTitleCWT","Title","Mr")
							str="Anshuman"&i
							Call EnterText("edtFirstNameCWT","First Name",str)
							str="Ghosh"&i
							Call EnterText("edtLastNameCWT","Last Name",str)
							Call selectRadioButton("rdGenderCWT","Gender","Male")

'							If i=0  Then
'									Call EnterText("edtEmailCWT","Email","sharad.mali@abc.com")
'							Else
'									Call EnterText("edtEmailCWT","Email","sharad.mali@abc.com")
'							End If
							str="sharad.mali"&i&"@abc.com"
							Call EnterText("edtEmailCWT","Email",str) 
							Call EnterText("edtPrimaryNumberCWT","Primary Number","07425160478")
							Call EnterText("edtMobileCWT","Mobile","07425160478")

							Call SelectList("lstNationality","Nationality","British")
							Call SelectList("lstDualNationality","Dual Nationality","British")
							Call VerifyItems("lstCountryOfResidence","Country Of Residence","Verify")
							Call Entertext("edtOccupation","Occupation","Director")
													
							If gstrCurID=11 OR gstrCurID=23 OR gstrCurID=606 Then
				
							Else   													
										Call VerifyItems("lstRole","Role","Verify")
										Call selectItemInList("lstRole","Role","2")
							End If

							If gstrCurID=10 OR gstrCurID=11 OR gstrCurID=23 OR gstrCurID=33 OR gstrCurID=35 OR gstrCurID=79 OR gstrCurID=83 OR gstrCurID=88 OR gstrCurID=93 OR gstrCurID=98 OR gstrCurID=603 OR gstrCurID=604 OR  gstrCurID=158 OR gstrCurID=163 OR gstrCurID=168 OR gstrCurID=173Then ' Soletrader
				
							Else
								Call Entertext("edtPositionInCompany","Position In Company","Director")
							End If	


							Call setCheckBox("chkIndividualKAPCWT","Individual KAP CWT","ON")
							If i=0 Then
								Call setCheckBox("chkPrimaryContact","Primary Contact","ON")
							End If							
							Call setCheckBox("chkValidatedCWT","Validated","ON")
							Call setCheckBox("chkAuthorisedSignatory","Authorised Signatory","ON")
							Call ClickButton("btnSaveCWT","Save Individual","Click")	
					Next
			End If
		Else
				Call TakeScreenShot()
				gstrDesc =  "Page '" &strLabel & "' is not displayed."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11			
	End If

	IndividualDetailsCWT = bresult
End Function


'========================================================================================='
'Action For  IndividualDetailsRO
'========================================================================================='
Public Function IndividualDetailsRO(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objTable = gobjObjectClass.getObjectRef(strObject)		
    
	If objTable.Exist(1) Then
	
			If  strVal="UPDATE" Then 
						
    						Set oDesc = Description.Create() 
							oDesc("micclass").Value = "WebElement" 
							oDesc("class").Value = "expandRowDetails" 
							oDesc("html tag").Value = "A" 

							Set objElementCollection=objTable.ChildObjects(oDesc)
							NumberOfWebElements = objElementCollection.Count 

								For ntRow = 0 To NumberOfWebElements-1
'									For ntRow = 0 To 0

												wait 3
												Set objTable1 = gobjObjectClass.getObjectRef(strObject)
												Set oDesc1 = Description.Create() 
												oDesc1("micclass").Value = "WebElement" 
												oDesc1("class").Value = "expandRowDetails" 
												oDesc1("html tag").Value = "A" 

												Set objElementCollection1=objTable1.ChildObjects(oDesc)
												NumberOfWebElements1 = objElementCollection1.Count 

												objElementCollection1(ntRow).highlight							
												objElementCollection1(ntRow).FireEvent "Click"
												wait 3

'												If  ntRow=0 Then
'														Call EnterText("edtTitleRO","Title","Mr")
'												Else
'														Call EnterText("edtTitleRO","Title","Miss")
'												End If
'												
'												If ntRow=0  Then
'													Call EnterText("edtFirstNameRO","First Name","Sharad")
'												Else
'													Call EnterText("edtFirstNameRO","First Name","Dipti")
'												End If
'					
'												If ntRow=0  Then
'														Call EnterText("edtLastNameRO","Last Name","Mali")
'												Else
'														Call EnterText("edtLastNameRO","Last Name","Bhushetty")
'												End If
												Call selectRadioButton("rdGenderRO","Gender","Male")
												If ntRow=0 Then
																Call EnterText("edtEmailRO","Email","sharad.mali@abc.com")
												ElseIf ntRow=1 Then
																Call EnterText("edtEmailRO","Email","sharad.mali@abc.com")
												ElseIf ntRow=2 Then
																Call EnterText("edtEmailRO","Email","sharad.mali@abc.com")
												End If

												If ntRow=0 Then
																Call EnterText("edtMobileRO","Mobile","7438116882")
												ElseIf ntRow=1 Then
																Call EnterText("edtMobileRO","Mobile","7459114432")
												ElseIf ntRow=3 Then
																Call EnterText("edtMobileRO","Mobile","7459114432")
												End If

												If ntRow=0 Then
																Call EnterText("edtPrimaryNumberRO","Primary Number","7438116882")
												ElseIf ntRow=1 Then
																Call EnterText("edtPrimaryNumberRO","Primary Number","7459114432")
												ElseIf ntRow=2 Then
																Call EnterText("edtPrimaryNumberRO","Primary Number","7459114432")
												End If
												Call VerifyItems("lstRoleRO","Role","Verify")
												Call selectItemInList("lstRole","Role","2")
												Call Entertext("edtPositioninCompanyRO","Position In Company","Director")
												Call SelectList("lstNationalityRO","Nationality","British")
												Call SelectList("lstDualNationalityRO","Dual Nationality","British")

												Call setCheckBox("chkAuthorisedSignatoryRO","Authorised Signatory","ON")
												Call setCheckBox("chkValidatedRO","Validated","ON")	

												If ntRow=0 Then
																Call setCheckBox("chkPrimaryContactRO","Primary Contact","ON")
												End If
																							
												Call ClickButton("btnSaveIndiRO","Save Individual","Click")	
                               Next

			ElseIf  strVal="ADD" Then 

							For i=0 to 1
									Call ClickLink("lnkAddIndividualRO","Add Individual Link","Click")	
									wait 2
									If  i=0 Then
											Call EnterText("edtTitleRO","Title","Mr")
									Else
											Call EnterText("edtTitleRO","Title","Miss")
									End If
									
									If i=0  Then
										Call EnterText("edtFirstNameRO","First Name","Sharad")
									Else
										Call EnterText("edtFirstNameRO","First Name","Dipti")
									End If
		
									If i=0  Then
											Call EnterText("edtLastNameRO","Last Name","Mali")
									Else
											Call EnterText("edtLastNameRO","Last Name","Bhushetty")
									End If

									Call selectRadioButton("rdGenderRO","Gender","Male")

									If i=0  Then
											Call EnterText("edtEmailRO","Email","sharad.mali@abc.com")
									Else
											Call EnterText("edtEmailRO","Email","sharad.mali@abc.com")
									End If
		
									If i=0  Then
											Call EnterText("edtMobileRO","Mobile","7438116882")
									Else
											Call EnterText("edtMobileRO","Mobile","7459114432")
									End If
		
									If i=0  Then
											Call EnterText("edtPrimaryNumberRO","Primary Number","7438116882")
									Else
											Call EnterText("edtPrimaryNumberRO","Primary Number","7459114432")
									End If
									Call VerifyItems("lstRoleRO","Role","Verify")
									Call selectItemInList("lstRole","Role","2")
									Call EnterText("edtPositioninCompanyRO","Positionin Company","Director")
									Call SelectList("lstNationality","Nationality","British")								
									Call SelectList("lstDualNationality","Dual Nationality","British")

									Call setCheckBox("chkAuthorisedSignatoryRO","Authorised Signatory","ON")
									Call setCheckBox("chkValidatedRO","Validated","ON")
									Call setCheckBox("chkIndividualKAPRO","Individual KAP CWT","ON")
									If i=0 Then
										Call setCheckBox("chkPrimaryContactRO","Primary Contact","ON")
									End If							
																	
									Call ClickButton("btnSaveIndiRO","Save Individual","Click")	
						Next

			ElseIf strVal="VERIFY" Then 
						
    						Set oDesc = Description.Create() 
							oDesc("micclass").Value = "WebElement" 
							oDesc("class").Value = "expandRowDetails" 
							oDesc("html tag").Value = "A" 

							Set objElementCollection=objTable.ChildObjects(oDesc)
							NumberOfWebElements = objElementCollection.Count 

								For ntRow = 0 To NumberOfWebElements-1
												wait 3
												Set objTable1 = gobjObjectClass.getObjectRef(strObject)
												Set oDesc1 = Description.Create() 
												oDesc1("micclass").Value = "WebElement" 
												oDesc1("class").Value = "expandRowDetails" 
												oDesc1("html tag").Value = "A" 

												Set objElementCollection1=objTable1.ChildObjects(oDesc)
												NumberOfWebElements1 = objElementCollection1.Count 

												objElementCollection1(ntRow).highlight							
												objElementCollection1(ntRow).FireEvent "Click"
												wait 3
												Call VerifyDisplayProperty("edtTitleRO","Title","value")
												Call VerifyDisplayProperty("edtFirstNameRO","First Name","value")
												Call VerifyDisplayProperty("edtLastNameRO","Last Name","value")
												Call VerifyDisplayProperty("edtEmailRO","Email","value")
												Call VerifyDisplayProperty("edtMobileRO","Mobile","value")
												Call VerifyDisplayProperty("edtPrimaryNumberRO","Primary Number","value")
												Call VerifyDisplayProperty("chkPrimaryContactRO","Primary Contact","value")
												Call VerifyDisplayProperty("chkValidatedRO","Validated","value")
												Call ClickButton("btnSaveIndiRO","Save Individual","SaveIndividualClick")	
                               Next
			End If
		Else
				Call TakeScreenShot()
				gstrDesc =  "Page '" &strLabel & "' is not displayed."
				WriteHTMLResultLog gstrDesc, 0
				CreateReport  gstrDesc, 0		
				bResult = False
				objErr.Raise 11			
	End If

	IndividualDetailsRO = bresult
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:EducationDocumentAttached

'---------------------------------------------------------------------------------------------------------------------
Public Function EducationDocumentAttached(strVal)
	
	Dim bResult, objImage
	bResult=False	
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Call VerifySibling("EducationDocumentAttachedOne","Education Document Attached One","EducationDocumentAttachedGet")
	Call VerifySibling("EducationDocumentAttachedTwo","Education Document Attached Two","EducationDocumentAttachedGet")
	Call VerifySibling("EducationDocumentAttachedThree","Education Document Attached Three","EducationDocumentAttachedGet")
	Call VerifySibling("EducationDocumentAttachedFour","Education Document Attached Four","EducationDocumentAttachedGet")

End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:BusinessDetailsCWT

'---------------------------------------------------------------------------------------------------------------------
Public Function BusinessDetailsCWT(strVal)
	
	Dim bResult, objImage
	bResult=False	
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Call selectRadioButton("rdBusinessTypeUCA","Business Type","BusinessTypeUCASet")
	Wait 4
	Call Entertext("edtGroupNameUCA","Group Name","GroupNameUCASet")
	Call VerifyDisplayProperty("lstSchoolTypeUCA","School Type","SchoolTypeUCAGet")
'	Call selectRadioButton("rdIncomeComeFromActiveSources","Income Come From Active Sources","Yes")
	wait 2
'	Call Selectlist("lstBrandUCA","Brand","BrandUCASelect")
'	Call VerifyDisplayProperty("lstBrandUCA","Brand","BrandUCAGet")
'	Call Selectlist("lstCustomerSegmentUCA","Customer Segment","CustomerSegmentUCASelect")
'	Call VerifyDisplayProperty("lstCustomerSegmentUCA","Customer Segment","CustomerSegmentUCAGet")
	
End Function

'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:RequestedDocumentList

'---------------------------------------------------------------------------------------------------------------------
Public Function RequestedDocumentList(strVal)
	
	Dim bResult, objImage
	bResult=False	
	strVal = getTestDataValue(strVal)

	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	wait 5
	Call Clicklink("lnkAddDocument","Add Document","AddDocumentClick")
	wait 5	
	Call Selectlist("RequestedDocumentCategory","Requested Document Category","Legal Address")
	wait 5
	Call Selectlist("RequestedDocumentType","Requested DocumentsType","Solicitor's Letter")
	wait 5
	Call Entertext("RequestedDocumentEffectiveDate","RequestedDocumentEffectiveDate","FUTUREDATE")
	wait 5
	Call Selectlist("lstDocumentsStatus","Documents Status","Required")
	wait 5
'	Call Entertext("edtDocumentsInstructions","Documents Instructions","DocumentsInstructionsSet")
	Call Clickbutton("btnSaveRequestedDocumentList","Save Requested Document List","SaveRequestedDocumentListClick")
End Function
'---------------------------------------------------------------------------------------------------------------------
'Function Name    	:VerifyAuditLogs
'---------------------------------------------------------------------------------------------------------------------
Function VerifyAuditLogs(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	strArrDetails=Split(strObject,":")

	Set objTable = gobjObjectClass.getObjectRef(strArrDetails(0))		
	Set objLink = gobjObjectClass.getObjectRef(strArrDetails(1))    

				If  objTable.exist(1) Then
										If  objLink.exist(1) Then
													If IsObject(objLink) Then
																While objTable.GetRowWithCellText(strVal) < 0 And objLink.Exist(1)
																				objLink.Click
																				Wait(5)
																				arrTemp = Split(strObject,":")
																					Set objTable = gobjObjectClass.getObjectRef(strArrDetails(0))		
																					Set objLink = gobjObjectClass.getObjectRef(strArrDetails(1))	
																Wend
													End If
										End If
				
										If objTable.GetRowWithCellText(strVal) > 0 Then
																Set oDesc = Description.Create() 
																oDesc("micclass").Value = "WebElement" 
																Set objElementCollection = objTable.ChildObjects(oDesc)
																NumberOfWebElements = objElementCollection.Count
																For i = 0 To NumberOfWebElements - 1 
																	If Trim(objElementCollection (i).GetROProperty("innertext"))= Trim(strVal) Then 
																							objElementCollection(i).Highlight
																										nRows=objTable.GetRowWithCellText(strVal)
																										nCols=objTable.ColumnCount(1)
																										For k=1 to nCols

																												gstrDesc = objTable.GetCellData(1,k) & "  --> " &  objTable.GetCellData( nRows,k)
																												WriteHTMLResultLog gstrDesc, 1
																												CreateReport  gstrDesc, 1
																										Next	
																												gstrDesc = ""
																												WriteHTMLResultLog gstrDesc, 5
																												CreateReport  gstrDesc, 1

																							bResult = True
																							Exit for
																	End If
																Next 															
												
										Else
																bResult = False
												
										End If
				Else
						bResult = False
				End If

			   If bResult Then
							gstrDesc = "Audit Log :  <B>" & strVal & "</B>. are present in the "& strLabel &" Report."
							WriteHTMLResultLog gstrDesc, 1
							CreateReport  gstrDesc, 1
							Browser("brwReportsTable").Close
							bResult = True
			   Else
							Call TakeScreenShot()
							gstrDesc = "Audit Log :  <B>" & strVal & "</B>. are not  present in the "& strLabel &" Report."
							WriteHTMLResultLog gstrDesc, 0
							CreateReport  gstrDesc, 0		
							bResult = False
							objErr.Raise 11
			   End If
			
				wait 2
VerifyAuditLogs=bResult
End function


'========================================================================================='
'Action For  UATIndividualDetailsCWT
'========================================================================================='
Public Function UATIndividualDetailsCWT(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)
	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If
	
	Set objTable = gobjObjectClass.getObjectRef(strObject)		
    
	If objTable.Exist(1) Then
	
			If  strVal="UPDATE" Then 
						
    						Set oDesc = Description.Create() 
							oDesc("micclass").Value = "WebElement" 
							oDesc("class").Value = "expandRowDetails" 
							oDesc("html tag").Value = "A" 

							Set objElementCollection=objTable.ChildObjects(oDesc)
							NumberOfWebElements = objElementCollection.Count 

									For ntRow = 0 To 1
												wait 3
												Set objTable1 = gobjObjectClass.getObjectRef(strObject)
												Set oDesc1 = Description.Create() 
												oDesc1("micclass").Value = "WebElement" 
												oDesc1("class").Value = "expandRowDetails" 
												oDesc1("html tag").Value = "A" 

												Set objElementCollection1=objTable1.ChildObjects(oDesc)
												NumberOfWebElements1 = objElementCollection1.Count 

												objElementCollection1(ntRow).highlight							
												objElementCollection1(ntRow).FireEvent "Click"
												Wait 2
												Call Entertext("edtMiddleName","edtMiddleName","Woodcock")
												Call Clickbutton("btnSaveCWT","Save","Click")
												wait 2
												Call Entertext("edtNotes","Notes","Automation Testing Note")
												Call Entertext("edtEffortTime","EffortTime","1")
												Call Clickbutton("btnSubmitUAT","Submit","Click")

                               Next
						End If
				End If
End function

'========================================================================================='
'Action For  GetCompanyName
'========================================================================================='
Public Function GetCompanyName(strObject,strLabel,strVal)
	Dim objChkBox, bresult
	bresult=False

	strVal = getTestDataValue(strVal)	
	If UCase(strVal) = "SKIP" Or objErr.Number = 11 Then
		Exit Function
	End If

	Set objComp = gobjObjectClass.getObjectRef(strObject)		
    
	If  objComp.EXIST(gExistCount) Then

		gstrCompanyName= objComp.GetROProperty("Value")
        gstrDesc =  "Successfully get the Company name :<B> '" & gstrCompanyName & "'</B>."
		WriteHTMLResultLog gstrDesc, 1
		CreateReport  gstrDesc, 1
		bResult = true
	Else
		gstrDesc =  "Failed to get the Company name."
		WriteHTMLResultLog gstrDesc, 0
		CreateReport  gstrDesc, 0		
		bResult = False
		objErr.Raise 11
	End If
GetCompanyName=bResult
End Function
