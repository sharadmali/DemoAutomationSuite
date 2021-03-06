﻿
'-------------------------------------------------------------------------------------------------
'Library File Name  :   ORClass
'Author             :   Sharad Mali
'Created date       :   
'Description        :   This class file has functions for reading XML OR description
'-------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
'Class Name	    :   clsOR
'Description	    :	Used to read the XML file and return the object reference.
'				The DLL has to be registered before using this
'-------------------------------------------------------------------------------------------------
class clsOR
	'Declare class variables
	
	Dim objReader 		' For XML Reader Object
	Dim strORFileName
	Dim gObjDictOR
	'-------------------------------------------------------------------------------------------------
	'Function Name	    :   fileExist
	'Input Parameter    :   strFileName (String)    -  XMLfile name
	'Description        :	Checks whether the XML file exists or not
	'Calls              :	None
	'Return Value	    :	True/False
	'-------------------------------------------------------------------------------------------------
	private Function fileExist(strFileName)
		dim objTextFile			' For file object
		' Create a file object
		Set objTextFile = CreateObject("Scripting.FileSystemObject")			
		'Check whether the 
		if objTextFile.FileExists(strFileName) then
			fileExist=True
		else
			fileExist=False	
		end if
	End function	
	'-------------------------------------------------------------------------------------------------
	'Function Name	    :   Class_Initialize
	'Input Parameter    :   None
	'Description        :	This function will be called when creating an object of clsOR
	'			(Implicit Call)
	'Calls              :	None
	'Return Value	    :	None
	'-------------------------------------------------------------------------------------------------
	Private Sub Class_Initialize()
	    ' Set the class variable for Object Repository Util object	
	    Set objReader = CreateObject("Mercury.ObjectRepositoryUtil")
	End Sub

	'-------------------------------------------------------------------------------------------------
	'Function Name	    :   Class_Terminate
	'Input Parameter    :   None
	'Description        :	This function will be called when unsetting reference to clsOR object
	'			(Implicit Call)
	'Calls              :	None
	'Return Value	    :	None
	'-------------------------------------------------------------------------------------------------
	Private Sub Class_Terminate()
	    'unset the object reference
	    set objReader = Nothing
	End Sub

	'-------------------------------------------------------------------------------------------------
	'Function Name	    :   setORFile(strFileName)
	'Input Parameter    :   strFileName(String) - XML file name
	'Description        :	To load the XML OR for reading
	'Calls              :	None
	'Return Value	    :	None
	'-------------------------------------------------------------------------------------------------
	public Function setORFile(strFileName)
		Dim bRC, nCount
		Dim strORFileName, strTempFile, objFile, objFSO
        bRC = True
		strORFileName = strFileName
		strTempFile = gstrObjectRepositoryDir & "\Temp\Temp_"& gstrProjectUser &".tsr"
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.GetFile(strFileName)
		objFile.Copy strTempFile
		strFileName = strTempFile
		Set objFile = Nothing
		Set objFSO = Nothing
		
		If fileExist(strFileName)=true then
			strORFileName = strFileName
			On Error Resume next
			nCount = 0
			Do
				objErr.Clear
				objReader.Load (strFileName)
				nCount = nCount + 1
				If nCount > 5 Then
					Exit Do
				End If
			Loop While(objErr.Description <> "")
			On Error goto 0
			'Create global dictionary object			
			Set gObjDictOR = CreateObject("Scripting.Dictionary")												
			EnumerateAll NULL,""		
		Else
			msgbox "Check the TSR File"
			bRC = False
		End If
		setORFile=bRC 
	End Function

	'---------------------------------------------------------------------------------------------
	'Function Name	 : getObjectRef
	'Input Parameter : strName - Logical name of the object
	'Description     : Gets string representing the properties of the logical name passed
	'Calls           :
	'Return Value	 : Objects corresponding to the names passed
	'---------------------------------------------------------------------------------------------
	Public Function getObjectRef(strName)		
	
		Dim objReturn, strExec,strObjectPath
		Set objReturn = Nothing		
		
		strObjectPath = Trim(gobjDictOR.Item(strName))	
		Set getObjectRef = objReturn

		If strObjectPath = "" Then
			Msgbox "Object " & strName & " Not found in object repository"
		Else
			'strExec = "Set objReturn = " & strObjectPath
			
			Set gObjectpath= Eval(strObjectPath)
			Set getObjectRef = gObjectpath
            strClassName = gObjectpath.GetTOProperty("Class Name") 
			On Error Resume Next
			If Strcomp( strClassName,"WebRadioGroup",1) = 0 OR Strcomp( strClassName,"WebList",1) = 0 OR Strcomp( strClassName,"Link",1) = 0 OR Strcomp( strClassName,"WebElement",1) = 0 OR Strcomp( strClassName,"WebButton",1) = 0 OR Strcomp( strClassName,"WebEdit",1) = 0 OR Strcomp( strClassName,"WebCheckBox",1) = 0 OR  Strcomp( strClassName,"WebTable",1) = 0 OR Strcomp( strClassName,"Frame",1) = 0 Then
				gObjectpath.ParentSync
			ElseIf Strcomp( strClassName,"Page",1) = 0  then
				gObjectpath.Sync
			End If
			On Error GoTo 0
		End If
				

	End Function
	
	'---------------------------------------------------------------------------------------------
	'Function Name	 : EnumerateAll
	'Input Parameter : strName - Logical name of the object
	'Description     : Gets string representing the properties of the logical name passed
	'Calls           :
	'Return Value	 : Objects corresponding to the names passed
	'---------------------------------------------------------------------------------------------
	Function EnumerateAll(strTestObject, strPath) 

		Dim TOCollection, TestObject, i, strTemp, strLogicalName
    		Set TOCollection = objReader.GetChildren(strTestObject) 		
    		For i = 0 To TOCollection.Count - 1 

            		Set TestObject = TOCollection.Item(i)
			strLogicalName = objReader.GetLogicalName(TestObject)
			If strPath = ""  Then
				strTemp =  TestObject.GetTOProperty("micclass") & "(" & Chr(34) & strLogicalName & Chr(34) & ")"
			Else
				strTemp = strPath & "." & TestObject.GetTOProperty("micclass") & "(" & Chr(34) & strLogicalName & Chr(34) & ")"
			End If			
			If gobjDictOR.Exists(strLogicalName) Then
				Msgbox "Object '" & strLogicalName & "' is duplicate in OR File " & strORFileName				             			
			Else
				gobjDictOR.Add objReader.GetLogicalName(TestObject),strTemp				
			End If 												
			EnumerateAll TestObject,strTemp
    		Next 

	End Function 

    Public Function getWinObjectRef(strName)		
	
		Dim objReturn, strExec,strObjectPath
		Set objReturn = Nothing		
		
		strObjectPath = gobjDictOR.Item(strName)	
		
		If strObjectPath = "" Then
			Msgbox "Object " & strName & " Not found in object repository"
		Else
			strExec = "Set objReturn = " & strObjectPath
			
			gObjectpath=strObjectPath
			
			Execute strExec	'Execute the string in QTP and returns the reference to the object 
		End If
		
		Set getObjectRef = objReturn
End Function

End Class
