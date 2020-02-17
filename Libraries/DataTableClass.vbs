'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'Library File Name  :   DataTableClass
'Author             :   Sharad Mali
'Created date       :   
'Description        :   This class file has functions for manupulating Excel Files
'-------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Class Name     :   clsExcel
'Description      : Used to create Excel object, read the contents and return the data in an array
'     Assumed that the first row in Excel shee is the column header
'-------------------------------------------------------------------------------------------------

class clsExcel
  'Declare class variables

  dim objXLS      ' For Excel Object
  dim objSheet      ' For Sheet object
  dim nSheet      ' for Sheet id
  dim strSheetName    ' to get the sheet name
  '-------------------------------------------------------------------------------------------------
  'Function Name      :   fileExist
  'Input Parameter    :   strFileName (String)    -  Excel file name
  'Description        : Checks whether the Excel file exists or not
  'Calls              : None
  'Return Value     : True/False
  '-------------------------------------------------------------------------------------------------
  public Function fileExist(strFileName)
    dim objTextFile     ' For file object
    ' Create a file object
    Set objTextFile = CreateObject("Scripting.FileSystemObject")
    'Check whether the
    if objTextFile.FileExists(strFileName) then
      fileExist=True
    else
      fileExist=False
      Set objTextFile = Nothing
    end if
  End function
  '-------------------------------------------------------------------------------------------------
  'Function Name      :   Class_Initialize
  'Input Parameter    :   None
  'Description        : This function will be called when creating an object of clsExcel
  '     (Implicit Call)
  'Calls              : None
  'Return Value     : None
  '-------------------------------------------------------------------------------------------------
  Private Sub Class_Initialize()
      ' Set the class variable for Excel object
      set objXLS = CreateObject("Excel.Application")
  End Sub

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   Class_Terminate
  'Input Parameter    :   None
  'Description        : This function will be called when unsetting reference to clsExcel object
  '     (Implicit Call)
  'Calls              : None
  'Return Value     : None
  '-------------------------------------------------------------------------------------------------
  Private Sub Class_Terminate()
      'unset the object reference
      set objXLS = Nothing
  End Sub

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   setSheet(nSheetId)
  'Input Parameter    :   strFileName(String) - Excel file name
  'Description        : To return Number of rows in Excel Sheet. It is assumed that there is no blank
  '     rows in between.
  'Calls              : None
  'Return Value     : nRowCount(int) - number of rows
  '-------------------------------------------------------------------------------------------------
  public Function setSheet(nSheetId)
    nSheet=nSheetId
  End Function

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   getRowCount
  'Input Parameter    :   strFileName(String) - Excel file name
  'Description        : To return Number of rows in Excel Sheet. It is assumed that there is no blank
  '     rows in between.
  'Calls              : None
  'Return Value     : nRowCount(int) - number of rows
  '-------------------------------------------------------------------------------------------------
  public Function getRowCount(strFileName)

    ' Define local variables
    Dim nLoopCount, nRowCount
    nLoopCount=2    ' First row in the excel sheet represents the column names

    nRowCount=0
    ' Open the Excel file
    objXLS.Workbooks.Open(strFileName)
    ' Select the sheet in the workbook

    Set objSheet = objXLS.ActiveWorkbook.Worksheets(nSheet)
    'Loop until a blank row

    Do
      If trim (objSheet.Cells(nLoopCount, 1)) = "" then
        Exit do
      End if
      nLoopCount = nLoopCount+1
      nRowCount = nRowCount +1
    Loop

    objXLS.ActiveWorkbook.Close
    set objSheet=Nothing
    ' Return the row count
    getRowCount=nRowCount
  End function

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   getColumnCount
  'Input Parameter    :   strFileName(String) - Excel file name
  'Description        : To return Number of columns in Excel Sheet. It is assumed that there are no
  '     blank columns in between in 1st row of Excel Sheet.
  'Calls              : None
  'Return Value     : nColCount(int) - number of rows
  '-------------------------------------------------------------------------------------------------
  public Function getColumnCount(strFileName)
    ' Define local variables
    Dim nLoopCount,nColCount
    nColCount=0
    nLoopCount=1
    ' Open the Excel file
    objXLS.Workbooks.Open(strFileName)
    ' Select the sheet in the workbook
    Set objSheet = objXLS.ActiveWorkbook.Worksheets(1)
    'Loop until a blank column
    Do
      If trim (objSheet.Cells(1, nLoopCount)) = "" then
        Exit do
      End if
      nLoopCount = nLoopCount+1
      nColCount = nColCount +1
    Loop
    objXLS.ActiveWorkbook.Close
    set objSheet=Nothing
    ' Return the column count
    getColumnCount=nColCount
  End function

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   getXLSDetails
  'Input Parameter    :   strFileName(String) - Excel file name
  '     nColCount(int)  - Represents number of columns
  '     arrData(String array) - passed by reference representing data in Excel
  'Description        : Reads the excel file & populates data array.
  'Calls              : None
  'Return Value     : None
  '-------------------------------------------------------------------------------------------------
  public Function getXLSDetails(strFileName,nColCount ,ByRef arrData())
    ' Define local variables
    dim nUbound
    dim nLoopCntOuter
    dim nLoopCntInner
    'Get the upper bound of array
    nUbound= UBound(arrData,1)
    ' Open the Workbook
    objXLS.Workbooks.Open(strFileName)
    ' Open the sheet
    Set objSheet = objXLS.ActiveWorkbook.Worksheets(nSheet)
    strSheetName = objSheet.name  'get the sheet name

    ' Loop thru all the rows and populate data array
    for nLoopCntOuter= 0 to nUbound
      for nLoopCntInner=0 to nColCount-1
        arrData(nLoopCntOuter,nLoopCntInner)= objSheet.Cells(nLoopCntOuter+2, nLoopCntInner+1)
      Next
    Next
    objXLS.ActiveWorkbook.Close
    set objSheet=Nothing
  End Function

  '-------------------------------------------------------------------------------------------------
  'Function Name      :   getsheetName
  'Input Parameter    :
  'Description        : Returns the name of the sheet
  'Calls              : None
  'Return Value     : Sheet name
  '-------------------------------------------------------------------------------------------------
  public Function getSheetName
    getsheetName=strSheetName
  End function



  '-------------------------------------------------------------------------------------------------
  'Function Name      :   getFrameworkXLSDetails
  'Input Parameter    :   strFileName(String) - Excel file name
  '     arrData(String array) - passed by reference representing data in Excel
  'Description        : Reads the excel file & populates data array. This function is specific to
  '       Keyword driven framework
  'Calls              : None
  'Return Value     : None
  '-------------------------------------------------------------------------------------------------
  public Function getFrameworkXLSDetails(strFileName,ByRef arrData())
    ' Define local variables
    dim nUbound
    dim nLoopCntOuter
    dim nLoopCntInner
    'Get the upper bound of array
    nUbound= UBound(arrData,1)
    ' Open the Workbook
    objXLS.Workbooks.Open(strFileName)
    ' Open the sheet
    Set objSheet = objXLS.ActiveWorkbook.Worksheets(nSheet)
    strSheetName = objSheet.name  'get the sheet name
    ' Loop thru all the rows and populate data array
    for nLoopCntOuter= 0 to nUbound
      arrData(nLoopCntOuter,0)= objSheet.Cells(nLoopCntOuter+2, 1) ' For Keyword
      arrData(nLoopCntOuter,1)= objSheet.Cells(nLoopCntOuter+2, 2) ' For Object Name
      arrData(nLoopCntOuter,2)= objSheet.Cells(nLoopCntOuter+2, 3) ' For parameters
      arrData(nLoopCntOuter,3)= objSheet.Cells(nLoopCntOuter+2, 4) ' For remarks
    Next
    objXLS.ActiveWorkbook.Close
    set objSheet=Nothing
    End Function

End Class

'-------------------------------------------------------------------------------------------------
'Function Name      :   getXLSClassObject
'Input Parameter    :   None
'Description        : To instantiate clsExcel object
'Calls              : None
'Return Value     : clsExcel object
'-------------------------------------------------------------------------------------------------

Function getXLSClassObject
  Set getXLSClassObject = New clsExcel
End Function
