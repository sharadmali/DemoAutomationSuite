﻿
' As of now Ignore  this class. Assume that this works fine
'-------------------------------------------------------------------------------------------------
'Library File Name  	:InitScript
'Author             	:Sharad Mali
'Created date       	:
'Description        	:
'-------------------------------------------------------------------------------------------------

Class clsIteration
	Dim nRowCnt
	Dim nCoCnt
	Dim arrColNames()
	Dim arrData()	

	Public function getTestCaseData
		dim strQuery
		dim rs, fso, bAccessTestDataFlag
		dim nrow
		dim ncol
		dim conn
		Dim nIndex, DataValues 

		nrow = 0
		ncol = 0
		nIndex = 0
		Set fso = CreateObject("Scripting.FileSystemObject")
		
'		If fso.FileExists(gstrTestDataDir & "\TestData_"& gstrEnv &".mdb") Then
		If fso.FileExists(gstrTestDataDir & "\TestData.mdb") Then
			bAccessTestDataFlag = True
		ElseIf fso.FileExists(gstrTestDataDir & "\TestData.xls") Then
			bAccessTestDataFlag = False
		Else
			MsgBox "Test Data not found."
			Exit Function
		End if

		Set fso = Nothing

        set conn = CreateObject("ADODB.Connection")
		set rs = CreateObject("ADODB.Recordset")
		set objrs1 = CreateObject("ADODB.Recordset")
	'	Connect with the DB
		on error resume next		
		If bAccessTestDataFlag = True Then
			'Setup the required query to get the correct values from the database
			strQuery = "Select count(*) from [TestCase] where TestCase ='" & gstrCurScenario &"'"
'			conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData_"& gstrEnv &".mdb;User Id=;Password=;"
			conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrTestDataDir & "\TestData.mdb;User Id=;Password=;"
		Else
			'Setup the required query to get the correct values from the database
			strQuery = "Select count(*) from [TestCase$] where TestCase ='" & gstrCurScenario &"'"
			conn.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & gstrTestDataDir & "\TestData.xls"					
		End If
		
		rs.open strQuery,conn,1,1		
		nRowCnt = rs(0)			
		rs.close
		If bAccessTestDataFlag = True Then
			strQuery = "Select * from [TestCase] where TestCase ='" & gstrCurScenario  &"'"
		Else
			strQuery = "Select * from [TestCase$] where TestCase ='" & gstrCurScenario  &"'"
		End If
		
		rs.open strQuery,conn,1,1

		'Loop to read the recordset till the end
		nCoCnt =  rs.Fields.Count		
		redim arrColNames(nCoCnt-1)
		while(ncol < nCoCnt)			
			arrColNames(nCol)= rs.Fields(ncol).Name
			ncol = ncol + 1
		wend
		ncol = 0
		nIndex=0
		ReDim arrData(nRowCnt-1,nCoCnt-1)
		while(not rs.EOF)					
			'Loop to add the index and fieldname mapping to the dictionary object		
			while(ncol < nCoCnt)			
			'Add values to the dictionary object with the field name as the value and the index as the key
				arrData(nIndex,ncol) =rs.Fields(ncol).Value	
				ncol = ncol + 1
			wend
			ncol = 0
			nIndex = nIndex+1
			'Move to the next recordset
			rs.movenext				
		wend
		rs.close
		conn.close
		If instr (arrData(0,1),",") <> 0 Then
			DataValues = split(arrData(0,1),",")
			gDataCount = Ubound(DataValues)
		Else
			gDataCount = 0
		End If						
	End Function
End Class
