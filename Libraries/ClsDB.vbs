'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'Library File Name  :   ClsDB
'Author             :   Sharad Mali
'Created date       :   
'Description        :   This class file contains Database related functions
'-------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Class Name :   ClsDB
'Description  :   Used to create DB object, read the contents and return the data in an array
'-------------------------------------------------------------------------------------------------
Class ClsDB
  private conn
  private rs
  '-------------------------------------------------------------------------------------------------
  'Function Name      :   Class_Initialize
  'Input Parameter    :   None
  'Description        : This function will be called when creating an object of clsDB
  '           (Implicit Call)
  'Calls              : None
  'Return Value     : None
  '-------------------------------------------------------------------------------------------------
  Private Sub Class_Initialize()
      ' Set the class variable for ADODB object
    set conn = CreateObject("ADODB.Connection")
    set rs = CreateObject("ADODB.Recordset")

  End Sub

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   Class_Terminate
  'Input Parameter    :   None
  'Description        : This function will be called when unsetting reference to clsDB object
  '           (Implicit Call)
  'Calls              : None
  'Return Value     : None
  '-------------------------------------------------------------------------------------------------
  Private Sub Class_Terminate()
      'unset the object reference
    set conn = Nothing
    set rs = Nothing
  End Sub
  '-------------------------------------------------------------------------------------------------
  'Function Name      :   connectToDB
  'Input Parameter    :   strDBName(String) - Representing Database Path
  'Description        : To connect to Database
  'Calls              : ErrorHandler
  'Return Value     : True/False
  '-------------------------------------------------------------------------------------------------
  Public function FuncConnectToDB(strDBName)
    dim bResult
    bResult=true
    On Error Resume Next
   conn.Open "DRIVER={Microsoft Excel Driver (*.xls)};DBQ=" & "TestArtifacts\TestData\TestData.xls"
    If Err.number <> 0 then
      'Call Error handling routine here
      bResult=false
    End if
    connectToDB=bResult
  End Function
'-------------------------------------------------------------------------
  'Function Name      :   executeQuery
  'Input Parameter    :   strQuery(String) - Representing SQL Query
  '               arrData (String Array) - to store data - Passed by ref
  'Description        : To Execute the SQL Query and to populate the array
  'Calls    : ErrorHandler
  'Return Value    :  True/False
  '-------------------------------------------------------------------------------------------------

  Public function executeQuery(strQuery, ByRef arrData)

    Dim nRowCnt, nColCnt
    Dim nLoopCnt, nCnt
    nCnt=0
    rs.open strQuery,conn,1,1
    nRowCnt=rs.RecordCount
    nColCnt=rs.Fields.Count
    ReDim arrData(nRowCnt-1,nColCnt-1)
    Do while (NOT rs.EOF)
      for nLoopCnt=0 to nColCnt-1
        arrData(nCnt,nLoopCnt)=rs(nLoopCnt)
      next
      rs.movenext
      nCnt=nCnt+1
    Loop

  End Function
  '-------------------------------------------------------------------------------------------------
  'Function Name      :   updateQuery
  'Input Parameter    :   strQuery(String) - Representing SQL Query (Update Query)
  'Description        : To Execute the Update Query. At present it is used to store Trade ID & Date
  'Calls      : To Be done
  'Return Value    :  To BeDone
  '-------------------------------------------------------------------------------------------------
  Public function updateQuery(strUpdateQuery)
    on error resume next
    rs.Open strUpdateQuery, conn
  End Function

End Class
  '-------------------------------------------------------------------------------------------------
  'Function Name      :   disconnectDB
  'Input Parameter    :   None
  'Description        : To disconnect Database
  'Calls              : ErrorHandler
  'Return Value     : True/False
  '-------------------------------------------------------------------------------------------------
  Public function FuncDisconnectDB()
    dim bResult
    bResult=true
    On Error Resume Next
    rs.close
    conn.close
    If Err.number <> 0 then
      'Call Error handling routine here
      bResult=false
      Exit Function
    End if
  End Function
  '------------------------
