''--------------------------------- Object Sync Register Function ----
RegisterUserFunc "WebEdit","Sync","EditSync",True
RegisterUserFunc "WebButton","Sync","WebButtonSync",True
RegisterUserFunc "WebList","Sync","ListSync",True
RegisterUserFunc "WebTable","Sync","WebTableSync",True
RegisterUserFunc "WebCheckBox","Sync","CheckBoxSync",True
RegisterUserFunc "WebEdit","Sync","EditSync",True
RegisterUserFunc "Frame","Sync","FrameSync",True
RegisterUserFunc "Page","Sync","PageSync",True
RegisterUserFunc "Page","busyCheck","busyCheck",True
RegisterUserFunc "Frame","busyCheck","busyCheck",True

''---------------------------------- ParentSync Register Function
RegisterUserFunc "WebEdit","ParentSync","ParentSync",True
RegisterUserFunc "WebButton","ParentSync","ParentSync",True
RegisterUserFunc "WebList","ParentSync","ParentSync",True
RegisterUserFunc "WebTable","ParentSync","ParentSync",True
RegisterUserFunc "WebCheckBox","ParentSync","ParentSync",True
RegisterUserFunc "Frame","ParentSync","ParentSync",True
RegisterUserFunc "Link","ParentSync","ParentSync",True
RegisterUserFunc "WebElement","ParentSync","ParentSync",True
RegisterUserFunc "WebRadioGroup","ParentSync","ParentSync",True

''---------------------------------- Objects registered function
RegisterUserFunc "WebElement","Set","SetTextInWebElement",True
RegisterUserFunc "WebEdit","Set","SetTextInEditBox",True

''------------------------------------------------------------------------------------------------------------------------------------
Function ParentSync(Byval objObject)
   If objObject.exist(10) Then
		objObject.GetTOProperty("parent").Sync
	Else
		objErr.raise 11
   End If
End Function

' Synchronize Edit box
Function EditSync(Byval objEdit)
		If objEdit.Exist(20) = False Then
				Exit Function
		End If
        objEdit.WaitProperty "disabled",false,15000		
End Function

' Synchronize Frame Object
Function FrameSync(Byval objFrame)
		objFrame.GetTOProperty("parent").Sync
		objFrame.RefreshObject
		If objFrame.Exist(20) = False Then
				ObjErr.raise 11
				Exit Function
		End If

'		objFrame.WaitProperty "attribute/readyState","complete",15000

End Function



' Synchronize WebTable Object
Function WebTableSync(Byval objTbl)
		objTbl.GetTOProperty("parent").Sync
End Function

' Synchronize Page Object
Function PageSync(Byval objPage)
		objPage.GetTOProperty("parent").Sync

		If objPage.Exist(20) = False Then
				objErr.raise 11
				Exit Function
		End If
		objPage.RefreshObject
		objPage.Sync	

		strTimer = Timer
'		objPage.RefreshObject
'		Set objBusy = objPage.image("file name:=busyIndicator.gif","index:=0")
'		Do While(objBusy.exist(0) AND ((Timer - strTimer) < 60))
'				objBusy.RefreshObject
'				Print "Busy Image"
'				If strcomp(objBusy.Object.style.display,"",1) = 0  OR strcomp(objBusy.Object.style.display,"none",1) = 0  Then
'					Exit Do
'				End If
'				objPage.RefreshObject
'				Set objBusy = objPage.image("file name:=busyIndicator.gif","index:=0")
'        Loop
		
End Function


''---------------------------------------------------------------------------------------------------------------------
'' Function name : SetTextInWebElement
'' Description 		: This function is used to set text value in WebElement (shocked :-) )
'' Paramter 		: ObjEle : WebElement Object
''							strText : Value to be set in Webelement
''---------------------------------------------------------------------------------------------------------------------
Function SetTextInWebElement(Byval ObjEle, Byval strText)
		'strText = resolveDataValue(strText)
		ObjEle.Object.innertext = Trim(strText)
End Function


''---------------------------------------------------------------------------------------------------------------------
'' Function name : SetTextInEditBox
'' Description 		: This function is used to set text value in WebEdit 
'' Paramter 		: ObjText : WebElement Object
''							strText : Value to be set in Webelement
''---------------------------------------------------------------------------------------------------------------------
Function SetTextInEditBox(Byval ObjText, Byval strText)
		 'strText = Trim(resolveDataValue(strText))
		 If ObjText.WaitProperty("disabled",0,15000) Then
				ObjText.Set strText
		 End If
End Function
