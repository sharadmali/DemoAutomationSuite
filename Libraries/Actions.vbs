'========================================================================================='
'Action For Launcing URLs
'========================================================================================='
Public Function ApplicationURL()

	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If

	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0

	Call getData("Table=ApplicationURL;Columns=*")
	While nIDIndex <= Ubound(arrData)
		strApplicationName = getTestDataValue("APPURL")
		gstrApplicationName = strApplicationName
		strAppURL= dictApplicationURL(strApplicationName)
		gstrApplicationURL=strAppURL

		If UCASE(getTestDataValue("APPURL")) <> "SKIP"Then
			'Call invokeMozillaBrowser(strAppURL)
			Call invokeBrowser(strAppURL)
		End If
	
		If objErr.number = 11 Then
				nIDIndex =  Ubound(arrData) + 1
			Else
				nIDIndex = nIDIndex + 1
		End If	
	Wend
End Function

Public Function LoginGuru99()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	Call getData("Table=LoginGuru99;Columns=*")
	nIDIndex = 0
	loadOR "Guru99.tsr"
	While nIDIndex <= Ubound(arrData)
'======================================================================================='
	If UCASE(getTestDataValue("Guru99LoginCheck")) <> "SKIP"Then
		Call verifypage("pgLogin","Guru99 Login","Guru99LoginCheck")
		Call Entertext("edtUid","User Name","UserNameSet")
		Call Entertext("edtPassword","Password","PasswordSet")
		Call Clickbutton("btnLOGIN","Log In","LogInClick")
		wait 3
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	'UnloadOR()
End Function
'========================================================================================='
'Action For  Login PEGA
'========================================================================================='
Public Function LoginPega()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=LoginPEGA;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	gstrUserNamePEGA=UCASE(getTestDataValue("UserNameSet"))

	If UCASE(getTestDataValue("UserNameSet")) = "EXTERNALQA" OR UCASE(getTestDataValue("UserNameSet")) = "MMEXTERNALQA" OR UCASE(getTestDataValue("UserNameSet")) = "RBBEXTERNALQA" Then

		'For i=0 to 9
					Browser("brwWelcomeToPegaRULES").Refresh
					Browser("brwWelcomeToPegaRULES").Sync
					wait 1
			'Next
	End If

	If UCASE(getTestDataValue("UserNameSet")) = "6485962" OR UCASE(getTestDataValue("UserNameSet")) = "8731053_CWT" Then

		'For i=0 to 4
					Browser("brwWelcomeToPegaRULES").Refresh
					Browser("brwWelcomeToPegaRULES").Sync
					wait 1
			'Next
	End If

	If UCASE(getTestDataValue("WelcomeToPegaRULESCheck")) <> "SKIP"Then
		Call verifypage("pgWelcomeToPegaRULES","Welcome To Pega RULES","WelcomeToPegaRULESCheck")
		Call Entertext("edtUserName","User Name","UserNameSet")
		Call Entertext("edtPassword","Password","PasswordSet")
		Call Clickbutton("btnLogIn","Log In","LogInClick")
		wait 3
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	'UnloadOR()
End Function

'========================================================================================='
'Action For  Create Case
'========================================================================================='
Public Function CreateCase()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CreateCase;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CreateCaseCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CreateCaseCheck")
		Call verifypage("pgPegaCaseManagerPortal","Create Case","CreateCaseCheck")
		Call ClickLink("lnkCreateCase","Create Case","CreateCaseClick")
		 Call VerifyProperty("elmSetupCaseTitle","Setup Case","EXIST=[TRUE]")
		Call GetCaseID("elmCaseID","CaseID","CaseIDGet")
		Call CaseID("UPDATE")
		Call selectRadioButton("rdSearchOrFindCompanyByID","Search Or Find Company ByID","SearchOrFindCompanyByIDSelect")

		Call Entertext("edtSearchCompanyName","Search Company Name","SearchCompanyNameSet")

		Call Entertext("edtCompanyID","Company ID","CompanyIDSet")
		Call Clickbutton("btnFindCompany","Find Company","FindCompanyClick")
		Wait 4
		Call GetCompanyName("edtCompanyNameCWT","Company Name","CompanyNameGet")
		Call selectDropdown("lstBrandUCA","Brand","BrandSelect")
		wait 2
		Call VerifyItems("lstBusinessUnit","Business Unit","BusinessUnitSelect")
		Call selectDropdown("lstBusinessUnit","Business Unit","BusinessUnitSelect")
		wait 2
		Call VerifyItems("lstSchoolType","School Type","SchoolTypeSelect")
		Call selectDropdown("lstSchoolType","School Type","SchoolTypeSelect")
		wait 2
		Call EnterPegaFormatDate("edtInitialClientMeeting","Initial Client Meeting","InitialClientMeetingSet")
		Call Entertext("edtRelationshipOwner","Relationship Owner","RelationshipOwnerSet")     
		wait 2
		Call PegaFooter("PegaFooterCheck")	
		wait 2
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Application Confirmation
'========================================================================================='
Public Function ApplicationConfirmation()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ApplicationConfirmation;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
    	Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkApplicationConfirmation(InitialResearch)","Application Confirmation (InitialResearch)","ApplicationConfirmationClick")	
    End If
	
	Call ClickLink("lnkApplicationConfirmation(InitialResearch)","Application Confirmation (InitialResearch)","ApplicationConfirmationDirectClick")	

	If UCASE(getTestDataValue("ApplicationConfirmationCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","ApplicationConfirmationCheck")
		Call verifypage("pgPegaCaseManagerPortal","Application Confirmation","ApplicationConfirmationCheck")
        Call VerifyProperty("elmApplicationConfirmationTitle","Application Confirmation Title","EXIST=[TRUE]")
		Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","ACUpdateClientsApplicationClick")	
		Call setCheckBox("chkROConfirms","ROConfirms","ROConfirmsSet")
		wait 5
		Call Entertext("edtNotes","Notes","NotesSet")
		Call Entertext("edtEffortTime","EffortTime","EffortTimeSet")
		Call Clickbutton("btnSubmit","Submit","SubmitClick")

		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Gather External Data
'========================================================================================='
Public Function GatherExternalData()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=GatherExternalData;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
    	Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkGatherExternalData","Gather External Data","GatherExternalDataClick")
	End If

    Call ClickLink("lnkGatherExternalData","Gather External Data","GatherExternalDataDirectClick")
	

	If UCASE(getTestDataValue("GatherExternalDataCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","GatherExternalDataCheck")
		Call verifypage("pgPegaCaseManagerPortal","Gather External Data","GatherExternalDataCheck")
        Call VerifyProperty("elmGatherExternalDataTitle","Gather External Data","EXIST=[TRUE]")
		Call Clicklink("lnkCharityRegisteredWebPage","Charity Registered WebPage","CharityRegisteredWebPageClick")
		Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","GEDUpdateClientsApplicationClick")	

		If gstrCurID="100" OR gstrCurID="500"Then
				Call VerifyDisplayProperty("lstEntityType","Entity Type","EntityTypeVerify")
				Call VerifyItems("lstEntityType","Entity Type","EntityTypeSelect")
				Call selectList("lstEntityType","Entity Type","EntityTypeSelect")
		Else			
				Call VerifyItems("lstCustomerTypeCWT","Customer Type CWT","CustomerTypeCWTVerify")
				gstrCustomerType=getTestDataValue("CustomerTypeCWTSelect")
				Call selectDropdown("lstCustomerTypeCWT","Customer Type CWT","CustomerTypeCWTSelect")
				Wait 5
				Call VerifyItems("lstCustomerSubTypeCWT","Customer Sub Type CWT","CustomerSubTypeCWTVerify")
				Call selectDropdown("lstCustomerSubTypeCWT","Customer Sub Type CWT","CustomerSubTypeCWTSelect")
				Wait 5
		End If
		
		Call EnterText("edtOnboardingManager","Onboarding Manager","OnboardingManagerSet")	
        Call PegaFooter("PegaFooterCheck")

		If UCASE(getTestDataValue("CloseCaseClick")) <> "SKIP"Then
				Call Clickbutton("btnClose","Close Case","CloseCaseClick")
				If Browser("brwOtherActions").Page("pgOtherActions").Frame("Page---->OtherActions").WebButton("btnDiscardOA").Exist(2) Then
						Browser("brwOtherActions").Page("pgOtherActions").Frame("Page---->OtherActions").WebButton("btnDiscardOA").Click
				End If
		 End If
		
	End If

	If UCASE(getTestDataValue("CharityRegisteredWebCheck")) <> "SKIP"Then
		Call verifypage("pgCharityRegisteredWeb","Charity Registered Web","CharityRegisteredWebCheck")
		Call VerifyWebElementInPage("pgCharityRegisteredWeb","Charity Registered Number","CharityRegisteredWebVerify")
		Call closeBrowser("brwCharityRegisteredWeb","Charity Registered Web","Close")
	End if

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function


'========================================================================================='
'Action For  CheckOwnershipStructure
'========================================================================================='
Public Function CheckOwnershipStructure()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CheckOwnershipStructure;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
    	Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkCheckOwnershipStructure","Check Ownership Structure","CheckOwnershipStructureClick")
	End If

	If UCASE(getTestDataValue("CheckOwnershipStructureCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CheckOwnershipStructureCheck")
		Call verifypage("pgPegaCaseManagerPortal","Check Ownership Structure","CheckOwnershipStructureCheck")
		Call VerifyProperty("elmCheckOwnershipStructureTitle","Check Ownership Structure","EXIST=[TRUE]")
		Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","COSUpdateClientsApplicationClick")	
        Call PegaFooter("PegaFooterCheck")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Check Existing Customer
'========================================================================================='
Public Function CheckExistingCustomer()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CheckExistingCustomer;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
    	  Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkCheckExistingCustomer","Check Existing Customer","CheckExistingCustomerClick")
	End If

	If UCASE(getTestDataValue("CheckExistingCustomerCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CheckExistingCustomerCheck")
		Call verifypage("pgPegaCaseManagerPortal","Check Existing Customer","CheckExistingCustomerCheck")
		Call VerifyProperty("elmCheckExistingCustomerTitle","Check Existing Customer","EXIST=[TRUE]")
		Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","CECUpdateClientsApplicationClick")	
'		Call setCheckBox("chkExistingCustomer","Existing Customer","ExistingCustomerSet")
'		Call Entertext("edtSortCode","Sort Code","SortCodeSet")
'        Call Entertext("edtAccountNumber","Account Number","AccountNumberSet")
'		strExist=CheckObjectExist("lstID&VType","ID&V Type","ID&VTypeObjectVeriy")
'		If  strExist Then
'				Call SelectDropdown("lstID&VType","ID&V Type","ID&VTypeSet")
'                Call VerifyItems("lstID&VType","ID&V Type","ID&VTypeSet")
'		End If
'		Call VerifyDisplayProperty("chkID&VCompleted","ID&V Completed","ID&VCompletedVerify")
'		Call selectDropdown("lstPCMPIDVType1","ID&V Type","Identification documents")
'		Call selectDropdown("lstPCMPIDVType2","D&V Type","Identification documents")
'		wait 5
'		Set pg=Browser("Pega Case Manager Portal").Page("Pega Case Manager Portal_2")
'		Set WedtAccNumber=description.Create
'		WedtAccNumber("micclass").value="WebEdit"
'		WedtAccNumber("html id").value="IDVValidityDate"
''		WedtAccNumber("html tag").value="SELECT"
'		set edtcnt=pg.ChildObjects(WedtAccNumber)
'		For i= 0 to edtcnt.count-1
'			edtcnt(i).set date+30
'		Next
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'CRA
'========================================================================================='
Public Function CRA()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CRA;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkCaptureCRAInputs","Capture CRA Inputs","CaptureCRAInputsClick")
	End If

	If UCASE(getTestDataValue("CaptureCRAInputsCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CaptureCRAInputsCheck")
		Call verifypage("pgPegaCaseManagerPortal","Capture Initial CRA Inputs","CaptureCRAInputsCheck")

		Call ClickLink("lnkPCMPCaptureCRAInputs","Capture CRA Inputs","Click")        'Newly added to sprint 35
		wait 3
		Call VerifyProperty("elmCRATitle","Initial CRA Title","EXIST=[TRUE]")
		Call setCheckBox("chkProducts","Products","ProductsSet")
		Call selectRadioButton("rdIndustryList","Industry List","IndustryListSelect")

		If gstrCurID=606 or gstrCurID=203   Then	

		Else
			Call selectList("lstCountryOfIncorporation","Country Of Incorporation","CountryOfIncorporationSelect")
			Call VerifyItems("lstCountryOfIncorporation","Country Of Incorporation","CountryOfIncorporationSelect")
		End If

		Call selectList("lstDomicileCountry","Domicile Country","DomicileCountrySelect")
'		Call VerifyItems("lstDomicileCountry","Domicile Country","DomicileCountrySelect")
		Call selectRadioButton("rdDeliveryChannel","Delivery Channel","DeliveryChannelSelect")
		Call selectList("lstCRARating","CRA Rating","CRARatingSelect")
		Call VerifyItems("lstCRARating","CRA Rating","CRARatingSelect")
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function


'========================================================================================='
'ReferVHighCase InitialCRA
'========================================================================================='
Public Function ReferVHighCaseInitialCRA()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ReferVHighCaseInitialCRA;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	Call ClickLink("lnkReferCasesVeryHigh","Refer Cases Very High","ReferCasesVeryHighClick")

	If UCASE(getTestDataValue("ReferVHighCasesCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Refer V High Cases Initial CRA","ReferVHighCasesCheck")
		Call selectRadioButton("rdReferVHighCases","Refer V High Cases","ReferVHighCasesSelect")
       Call PegaFooter("PegaFooterCheck")
	   Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Clarification call initial review
'========================================================================================='
Public Function ClarificationCallInitialReview()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ClarificationCallInitialReview;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkIsClarificationCallRequired","Clarification Call Required","IsClarificationCallRequiredClick")
	End If
	Call ClickLink("lnkIsClarificationCallRequired","Clarification Call Required","IsClarificationCallRequiredDirectClick")

	If UCASE(getTestDataValue("ClarificationCallInitialReviewCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","ClarificationCallInitialReviewCheck")
		Call verifypage("pgPegaCaseManagerPortal","Clarification Call Initial Review","ClarificationCallInitialReviewCheck")
		Call VerifyProperty("elmClarificationCallInitialReview","Clarification Call Initial Review Title","EXIST=[TRUE]")
        Call Entertext("edtEffortTime","EffortTime","1")
		Call Clickbutton("btnSubmit","Submit","Click")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Capture Customer Info
'========================================================================================='
Public Function CaptureCustomerInfo()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CaptureCustomerInfo;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		'Call VerifyLabel("pgPegaCaseManagerPortal","Assigned To","AssignedToLabelVerify")
		str=Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebElement("elmPCMPSuspendedTasks").GetROProperty("innertext")
		If Trim(str)="SuspendedTasks" Then
			gstrDesc =  "<B>" & Trim(str) & " </B> Expexted labels for " & " Assigned To"
			WriteHTMLResultLog gstrDesc, 1
			CreateReport  gstrDesc, 1
		End If
		Call ClickLink("lnkCaptureCustomerInfo","Capture Customer Info","CaptureCustomerInfoClick")
	End If

	Call ClickLink("lnkCaptureCustomerInfo","Capture Customer Info","CaptureCustomerInfoDirectClick")

	If UCASE(getTestDataValue("CaptureCustomerInfoCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CaptureCustomerInfoCheck")
		Call verifypage("pgPegaCaseManagerPortal","Capture Customer Info","CaptureCustomerInfoCheck")
		Call VerifyProperty("elmCaptureCustomerInfo","Capture Customer Info Title","EXIST=[TRUE]")
		Call selectRadioButton("rdCustomerAvailable","Customer Available","CustomerAvailableSelect")
		wait 2
        Call PegaFooter("PegaFooterCheck")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Verfiy Customer Information
'========================================================================================='
Public Function VerfiyCustomerInformation()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=VerfiyCustomerInformation;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='

	If UCASE(getTestDataValue("VerfiyCustomerInformationCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","VerfiyCustomerInformationCheck")
		Call verifypage("pgPegaCaseManagerPortal","Verfiy Customer Information","VerfiyCustomerInformationCheck")
		'Call VerifyProperty("elmVerfiyCustomerInformation","Verfiy Customer Information Title","EXIST=[TRUE]")
		Call selectRadioButton("rdCustomerCallSecurityQuestions1","Customer Call Security Questions1","CustomerCallSecurityQuestions1Select")
		wait 5
		Call selectRadioButton("rdCustomerCallSecurityQuestions2","Customer Call Security Questions2","CustomerCallSecurityQuestions2Select")
		Wait 5
		Call Entertext("edtNotes","Notes","Automation Testing Note")
        Call Clickbutton("btnSubmit","Submit","Click")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Business Details
'========================================================================================='
Public Function NewUpdateClientsApplication()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=NewUpdateClientsApplication;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("AccountSetupCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Account Setup","BusinessDetailsCheck")
		Call VerifyProperty("elmPCMPAccountSetupDetails","Account Set upTitle","EXIST=[TRUE]")
		Call selectList("lstROArea","Area","Business Development Manager")
		Call selectRadioButton("rdBusinessType","Business Type","Startup")
		Call selectList("lstPCMPCustomerSegment","Customer Segment","CustomerSegment")
		'Call ClickWebElement("elmUpdateApplicationBusinessYes","Yes","Click")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
	If UCASE(getTestDataValue("BusinessDetailsCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Business Details","BusinessDetailsCheck")
		Call VerifyProperty("elmBusinessDetailsTitle","Business Details Title","EXIST=[TRUE]")
'		Call BusinessDetailsCWT("BusinessDetailsCWTVerify")
		If gstrCustomerType <>"Sole Trader"  Then
			Call selectRadioButton("rdIncomeComeFromActiveSources","Income Come From Active Sources","Yes")
		End If
'		Call ClickWebElement("elmUpdateApplicationBusinessYes","Yes","Click")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
	If UCASE(getTestDataValue("BusinessAddressCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Business Address","BusinessDetailsCheck")
		Call VerifyProperty("elmPCMPBusinessAddresses","Business Address Title","EXIST=[TRUE]")
		Call selectList("lstPCMPBusinessPremise","Business Premise","Owned")
		Call selectList("lstPCMPTradingPremise","Trading Premise","Office")
'		Call BusinessDetailsCWT("BusinessDetailsCWTVerify")
'		Call ClickWebElement("elmUpdateApplicationBusinessYes","Yes","Click")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
	If UCASE(getTestDataValue("IndividualDetailsCWTCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","Individual Details CWT","IndividualDetailsCWTCheck")
		Call VerifyProperty("elmIndividualsTitle","Individuals Title","EXIST=[TRUE]")
		Call IndividualDetailsCWT("wtbIndividualsDetails","Individuals Details","IndividualsDetailsVerify")
		Call IndividualDetailsCWT("wtbIndividualsDetails","Individuals Details","IndividualsDetailsVerify1")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
	If UCASE(getTestDataValue("KnowYourBusinessCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","KnowYourBusiness","NatureOfBusinessCheck")
		Call VerifyProperty("elmKnowYourBusiness","KnowYourBusiness Title","EXIST=[TRUE]")
'		Call Entertext("edtNatureOfBusinessCWT","edtNatureOfBusinessCWT","Business and management consultancy activities not elsewhere classified")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
'	If UCASE(getTestDataValue("NatureOfBusinessCheck")) <> "SKIP" Then
'		Call verifypage("pgPegaCaseManagerPortal","Nature Of Business","NatureOfBusinessCheck")
'		Call VerifyProperty("elmNatureOfBusinessTitle","Nature Of Business Title","EXIST=[TRUE]")
'		Call Entertext("edtNatureOfBusinessCWT","edtNatureOfBusinessCWT","Business and management consultancy activities not elsewhere classified")
'		Call Clickbutton("btnUCANext","Next","Click")
'	End If

	If UCASE(getTestDataValue("TurnoverDetailsCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","TurnoverDetails","TurnoverDetailsCheck")
		Call VerifyProperty("elmTurnOverDetails","TurnOverDetailsTitle","EXIST=[TRUE]")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
'	If UCASE(getTestDataValue("SourceOfFundsCheck")) <> "SKIP" Then
'		Call verifypage("pgPegaCaseManagerPortal","Source Of Funds","SourceOfFundsCheck")
'		Call VerifyProperty("elmSourceOfFundsTitle","Source Of Funds Title","EXIST=[TRUE]")
'		Call Clickbutton("btnUCANext","Next","Click")
'	End If
'	
'	If UCASE(getTestDataValue("PurposeOfAccountCheck")) <> "SKIP" Then
'		Call verifypage("pgPegaCaseManagerPortal","Purpose Of Account","PurposeOfAccountCheck")
'		Call VerifyProperty("elmPurposeOfAccountTitle","Purpose Of Account Title","EXIST=[TRUE]")
'		Call Clickbutton("btnUCANext","Next","Click")
'	End If
	If UCASE(getTestDataValue("ProductDetailsCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","Product Details","ProductDetailsCheck")
		Call VerifyProperty("elmProductDetailsTitle","Product Details Title","EXIST=[TRUE]")
		Call ProductsSelectionCWT("pgPegaCaseManagerPortal","Products Selecttion","ProductsSelecttionClick")
		Call Clickbutton("btnUCANext","Next","Click")
	End If

	If UCASE(getTestDataValue("CommunicationPreferencesCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","Communication Preferences","CommunicationPreferencesCheck")
		Call VerifyProperty("elmCommunicationPreferencesTitle","Communication Preferences Title","EXIST=[TRUE]")
		Call selectAllRadioButton("pgPegaCaseManagerPortal","Marketting Preference","Yes")
		Call Clickbutton("btnUCANext","Next","Click")
	End If

	If UCASE(getTestDataValue("RequiredDocListCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","Required Doc List","RequiredDocListCheck")
		Call VerifyProperty("RequiredDocListTitle","Required Doc List Title","EXIST=[TRUE]")
		Call Clickbutton("btnUCANext","Next","Click")
	End If
	If UCASE(getTestDataValue("CustomerPlaybackCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","Customer Playback","CustomerPlaybackCheck")
		Call setCheckBox("chkPCMPCustomerPlayback","Customer Playback","ON")
		'Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnUCANext","Next","Click")
		wait 2
		Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").Frame("Page -----> Update Client's Application").WebButton("btnUCANext").Click
	End If

	If UCASE(getTestDataValue("ClarificationCallConfirmationCheck")) <> "SKIP" Then
		Call verifypage("pgPegaCaseManagerPortal","Clarification Call Confirmation","ClarificationCallConfirmationCheck")
		Call VerifyProperty("elmClarificationCallConfirmationTitle","Clarification Call Confirmation Title","EXIST=[TRUE]")
		wait 2
		Call selectRadioButton("rdCallComplete","Call Complete","CallCompleteSelect")
		wait 2
		Call selectRadioButton("rdPaymentLimit","Payment Limit","false")             'Sprint 36
		Call Entertext("edtEffortTime","EffortTime","1")
		wait 5
		Call fireEvent("btnFinishUCA","Finish","Click")
	End If

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Check Existing Customer-ID&V
'========================================================================================='
Public Function CheckExistingCustomerIDV()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
		Call getData("Table=CheckExistingCustomerIDV;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")	
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkIdentificationAndVerification","Identification And Verification","IdentificationAndVerificationClick")
	End If

	If UCASE(getTestDataValue("IdentificationAndVerificationCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","IdentificationAndVerificationCheck")
		Call verifypage("pgPegaCaseManagerPortal","Identification And Verification","IdentificationAndVerificationCheck")
		Call VerifyProperty("elmID&VTitle","ID&V Title","EXIST=[TRUE]")
'		 Call setCheckBox("chkID&VCompleted","ID&V Completed","ID&VCompletedSet")
		Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")

		If Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").Link("lnkCompanyRiskCheck").Exist(5) Then
					Call ClickLink("lnkCompanyRiskCheck","Company Risk Check","Click")
					wait 2
					Call syncbrowser("brwPegaCaseManagerPortal","Check")
					Call verifypage("pgPegaCaseManagerPortal","Company Risk Check","Check")
					Call VerifyProperty("elmRiskCheckCompanyTitle","Risk Check Company Title","EXIST=[TRUE]")
					Call RiskChecksComp("pgPegaCaseManagerPortal","Risk Checks KAP","Verify")
					Call Entertext("edtEffortTime","EffortTime","1")
					Call Clickbutton("btnSubmit","Submit","Click")
		End If

		If Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").Link("lnkPCMPIndividualRiskChecks").Exist(5) Then
			Call ClickLink("lnkPCMPIndividualRiskChecks","Individual Risk Check","Click")
			wait 2
			Call syncbrowser("brwPegaCaseManagerPortal","Check")
			Call verifypage("pgPegaCaseManagerPortal","Risk Checks KAPs","Check")
			Call VerifyProperty("elmRiskChecksKAPs","Risk Checks KAPs","EXIST=[TRUE]")
			Call RiskChecksKAP("pgPegaCaseManagerPortal","Risk Checks KAP","Verify")
			Call Entertext("edtEffortTime","EffortTime","1")
			Call Clickbutton("btnSubmit","Submit","Click")
		End If

	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Action For  Risk Check Company_New
'========================================================================================='
Public Function RiskCheckCompany()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=RiskCheckCompany;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkCompanyRiskCheck","Company Risk Company","CompanyRiskCheckClick")
	End If

	If UCASE(getTestDataValue("CompanyRiskCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CompanyRiskCheck")
		Call verifypage("pgPegaCaseManagerPortal","Company Risk Check","CompanyRiskCheck")
		Call VerifyProperty("elmRiskCheckCompanyTitle","Risk Check Company Title","EXIST=[TRUE]")
		Call RiskChecksComp("pgPegaCaseManagerPortal","Risk Checks KAP","RiskChecksCompVerify")
		wait 5
       Call PegaFooter("PegaFooterCheck")
	   Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	 
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Risk Checks Individuals_New
'========================================================================================='
Public Function RiskChecksIndividuals()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=RiskChecksIndividuals;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
        Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
    	Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkIndividualRiskChecks","Individual Risk Checks","IndividualRiskChecksClick")
	End If

	If UCASE(getTestDataValue("IndividualRiskChecksCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","IndividualRiskChecksCheck")
		Call verifypage("pgPegaCaseManagerPortal","Risk Checks KAPs","IndividualRiskChecksCheck")
		Call VerifyProperty("elmRiskChecksKAPs","Risk Checks KAPs","EXIST=[TRUE]")
		Call RiskChecksKAP("pgPegaCaseManagerPortal","Risk Checks KAP","RiskChecksKAPVerify")
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'CRA Next
'========================================================================================='
Public Function CRANext()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CRANext;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkCaptureCRAInputsCustomerEngagement","Capture CRA Inputs Customer Engagement","CaptureCRAInputNextClick")
	End If
	
	If UCASE(getTestDataValue("CaptureCRAInputsNextCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CaptureCRAInputsNextCheck")
		Call verifypage("pgPegaCaseManagerPortal","Capture CRA Inputs Customer Engagement Next","CaptureCRAInputsNextCheck")
        Call VerifyProperty("elmCRANextTitle","Final CRA Next Title","EXIST=[TRUE]")
		Call VerifyDisplayProperty("lstCRARating","CRA Rating","CRARatingVerify")
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'ReferVHighCaseFinalCRA
'========================================================================================='
Public Function ReferVHighCaseFinalCRA()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ReferVHighCaseFinalCRA;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	Call ClickLink("lnkReferCasesVeryHigh","Refer Cases Very High","ReferCasesVeryHighClick")

	If UCASE(getTestDataValue("ReferVHighCasesCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Refer V High Cases Final CRA","ReferVHighCasesCheck")
		Call VerifyProperty("elmReferVeryHighCasesTitle","Refer Very High Cases Final CRA Title","EXIST=[TRUE]")
		wait 5
		Call selectRadioButton("rdCRARating","Refer V High Cases","ReferVHighCasesSelect")
       Call PegaFooter("PegaFooterCheck")
	   Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Mandate Invite
'========================================================================================='
Public Function MandateInvite()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=MandateInvite;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("MandateInviteCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Mandate Invite","MandateInviteCheck")
        Call VerifyProperty("elmMandateInviteTitle","Mandate Invite Title","EXIST=[TRUE]")
		Call VerifyPrimaryContact("pgPegaCaseManagerPortal","Primary Contact Details","IndividualDetailsVerify")
		Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Configure Mandate Sent
'========================================================================================='
Public Function ConfigureMandateSent()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ConfigureMandateSent;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call CaseID("SELECT")
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")	
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
        Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkConfigureMandateCustomerEngagement","Configure Mandate Sent","ConfigureMandateSentClick")
	End If

	If UCASE(getTestDataValue("ConfigureMandateSentCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","ConfigureMandateSentCheck")
		Call verifypage("pgPegaCaseManagerPortal","Configure Mandate Customer Engagement","ConfigureMandateSentCheck")
        Call VerifyProperty("elmConfigureMandateSentTitle","Configure Mandate Sent Title","EXIST=[TRUE]")
		Call setAllCheckBoxMandateSent("pgPegaCaseManagerPortal","ConfigureMandateSent","ON")
		Call selectRadioButton("rgpPCMPSigningInstruction","Signing  Instruction","2 to sign")           				'Sprint 39 mandatory change
		Call selectRadioButton("rdPaymentLimitSelect","Payment Limit","false")													'Sprint 36 change
		wait 1
        Call selectRadioButton("rdMandateChannel","Mandate Channel","e-mandate")									'Sprint 36 change
		wait 1
		'Call selectRadioButton("rdPaperMandateType","Paper Mandate Typel","Post")									'Sprint 36 change

		Call EnterText("edtPCMPSortCode","SortCode","121212")
		Call EnterText("edtPCMPAccounNumber","Account Number","12121313")
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Confirm Mandate Received
'========================================================================================='
Public Function ConfirmMandateReceived()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ConfirmMandateReceived;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)		
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkConfirmMandateReceived","Confirm Mandate Received","ConfirmMandateReceivedClick")
	End If

	If UCASE(getTestDataValue("ConfirmMandateReceivedCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","ConfirmMandateReceivedCheck")
		Call verifypage("pgPegaCaseManagerPortal","Confirm Mandate Received","ConfirmMandateReceivedCheck")
        Call VerifyProperty("elmConfirmMandateReceivedTitle","Confirm Mandate Received Title","EXIST=[TRUE]")
		Call setCheckBox("chkMandateReceived","Mandate Received","MandateReceivedSet")
		Call setAllCheckBoxMandateAccepted("pgPegaCaseManagerPortal","Confirm Mandate Received","MandateAcceptedSet")
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Prepare for CDD QC
'========================================================================================='
Public Function PrepareForCDDQC()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=PrepareForCDDQC;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)		
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkPrepareForCDDQC","Prepare For CDD QC","PrepareForCDDQCClick")
	End If

	If UCASE(getTestDataValue("PrepareForCDDQCCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","PrepareForCDDQCCheck")
		Call verifypage("pgPegaCaseManagerPortal","Prepare For CDD QC","PrepareForCDDQCCheck")
		Call VerifyProperty("elmPrepareForCDDQCTitle","Prepare For CDD QC Title","EXIST=[TRUE]")
		Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","CDDUpdateClientsApplicationClick")	
		Call UATIndividualDetailsCWT("wtbIndividualsDetails","Individuals Details","UATIndividualDetailsCWTCheck")
		Call selectRadioButton("rdOutstandingDocuments","Outstanding Documents","OutstandingDocumentsSelect")
		wait 4
       Call PegaFooter("PegaFooterCheck")
	   Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Update Clients Application
'========================================================================================='
Public Function UpdateClientsApplication()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=UpdateClientsApplication;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'=======================================================================================
    If UCASE(getTestDataValue("UpdateClientsApplicationCheck")) <> "SKIP"Then
		Wait 3
		Call syncbrowser("brwPegaCaseManagerPortal","UpdateClientsApplicationCheck")
		Call verifypage("pgPegaCaseManagerPortal","Update Clients Application","UpdateClientsApplicationCheck")
		Call VerifyLabel("pgPegaCaseManagerPortal","Account Setup Details","AccountSetupDetailsCWTLabelVerify")
		Call VerifyLabel("pgPegaCaseManagerPortal","Business Details","BusinessDetailsCWTLabelVerify")
		Call VerifyLabel("pgPegaCaseManagerPortal","Business addresses","BusinessAddressesCWTLabelVerify")

		Call ClickWebelement("elmPCMPIndividualExpand","Individual","Click")
		Call VerifyLabel("pgPegaCaseManagerPortal","Individuals","IndividualsCWTLabelVerify")

		Call VerifyLabel("pgPegaCaseManagerPortal","Ownership Structure","OwnershipStructureLabelVerify")
		Call VerifyLabel("pgPegaCaseManagerPortal","Know Your Business","KnowYourBusinessLabelVerify")
		Call VerifyLabel("pgPegaCaseManagerPortal","Turn Over Details","TurnOverDetailsLabelVerify")
		'Call VerifyLabel("pgPegaCaseManagerPortal","Source Of Funds","SourceOfFundsLabelverify")
		Call VerifyLabel("pgPegaCaseManagerPortal","Communication Preferences","CommunicationPreferencesLabelVerify")
		Call BusinessDetailsCWT("BusinessDetailsCWTVerify")
		Wait 5
		Call IndividualDetailsCWT("wtbIndividualsDetails","Individuals Details","IndividualsDetailsVerify")
		Call VerifyProductsSelecttion("pgPegaCaseManagerPortal","Products Selecttion","ProductsSelecttionVerify")
		Call ProductsSelectionCWT("pgPegaCaseManagerPortal","Products Selecttion","ProductsSelecttionClick")      
		Call ProductseDeletionCWT("pgPegaCaseManagerPortal","Products Deletion","ProductsDeletionClick")
		Call EducationDocumentAttached("EducationDocumentAttachedVerify")
		Call RequestedDocumentList("RequestedDocumentListCheck")
        Call Entertext("edtNotes","Notes","NoteSet")
		Call Entertext("edtEffortTime","Effort Time","EffortTimeSet")
		Call Clickbutton("btnSubmit_2","Submit","SubmitClick")
		wait 5
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")        
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function

'========================================================================================='
'Logout Pega
'========================================================================================='
Public Function LogoutPega()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=LogoutPEGA;Columns=*")
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("LogoutPegaCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","LogoutPegaCheck")
		Call verifypage("pgPEGALogout","Logout Pega","LogoutPegaCheck")
		Call CloseAllBrowsers("Check")
'		Call Clicklink("lnkLoutOutUsers","Lout Out Users","LoutOutUsersClick")
'		Call ClickWebelement("elmLoutOut","Lout Out","LoutOutClick")
   	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function












'========================================================================================='
'Action For  RiskCheckCompany_Sprint24
'========================================================================================='
Public Function RiskCheckCompany_S24()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=RiskCheckCompany;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkCompanyRiskCheck","Company Risk Company","CompanyRiskCheckClick")
	End If


	If UCASE(getTestDataValue("CompanyRiskCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CompanyRiskCheck")
		Call verifypage("pgPegaCaseManagerPortal","Company Risk Check","CompanyRiskCheck")
'		Call VerifyProperty("elmAverseMediaHits","Averse Media Hits","EXIST=[AverseMediaHitsVerify]")
		Call Clickwebelement("AdverseMediaNoFound","AdverseMediaNoFound","Click")
   		Call selectRadioButton("rdSanctions","Sanctions","SanctionsSelect")
		Call selectRadioButton("rdPEP","PEP","PEPSelect")
       Call PegaFooter("PegaFooterCheck")
	   Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function


'========================================================================================='
'Clarification Call
'========================================================================================='
Public Function ClarificationCall()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ClarificationCall;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
		Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkClarificationCall","Clarification Call","ClarificationCallClick")
	End If

	If UCASE(getTestDataValue("ClarificationCallCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","ClarificationCallCheck")
		Call verifypage("pgPegaCaseManagerPortal","Clarification Call","ClarificationCallCheck")
		Call VerifyDisplayProperty("chkInviteToEmandate","Invite To E-mandate","InviteToEmandateVerify")
		Call setCheckBox("chkInviteToEmandate","Invite To E-mandate","InviteToEmandateSet")
		Call VerifyPrimaryContact("pgPegaCaseManagerPortal","Primary Contact Details","IndividualDetailsVerify")
		Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","UpdateClientsApplicationClick")	

       Call PegaFooter("PegaFooterCheck")
	   Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function



'========================================================================================='
'RiskChecksIndividuals_Spriny24
'========================================================================================='
Public Function RiskChecksIndividuals_S24()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=RiskChecksIndividuals;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
        Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
    	Call fireEvent("elmSearch","Search","SearchClick")
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkIndividualRiskChecks","Individual Risk Checks","IndividualRiskChecksClick")
	End If

	If UCASE(getTestDataValue("IndividualRiskChecksCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","IndividualRiskChecksCheck")
		Call verifypage("pgPegaCaseManagerPortal","Identification And Verification","IndividualRiskChecksCheck")
		Call VerifyAdverseMedia("pgPegaCaseManagerPortal","Adverse Media","AdverseMediaVerify")
        Call setCheckBox("chkRiskChecksCompleted","Risk Checks Completed","ON")
        Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function




'========================================================================================='
'AMLRiskAssessment
'========================================================================================='
Public Function AMLRiskAssessment()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=AMLRiskAssessment;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgAMLAssessment","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)		
		Call fireEvent("elmSearch","Search","SearchClick")
'		Call VerifyProperty("elmCaseStatusAML","Case Status","EXIST=[TRUE]")
		Call clickWebElement("elmCDDAMLRiskAssessment","CDD AML Risk Assessment","CDDAMLRiskAssessmentClick")
		Call clickWebElement("elmCDDAMLRiskCheck","CDD AML Risk Check","CDDAMLRiskCheckClick")
		Call clickWebElement("elmOnboardingQA","Onboarding QA","OnboardingQAClick")
		Call Clicklink("lnkRepairQualityCheck","Repair Quality Check","RepairQualityCheckClick")
		Call clickWebElement("elmCloseMailCase","Close Mail Case","CloseMailCaseClick")
	End If

	If UCASE(getTestDataValue("AMLRiskAssessmentCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","AMLRiskAssessmentCheck")
		Call verifypage("pgAMLAssessment",getTestDataValue("AMLRiskAssessmentCheck"),"AMLRiskAssessmentCheck")
		Call Clicklink("lnkCDDAMLRiskAssessmentQualityAssurance","CDDAMLRiskAssessmentQualityAssurance","CDDAMLRiskAssessmentQualityAssuranceClick")
		Call ClickLink("lnkCDDAMLRiskQAQualityAssurance","CDD AML RiskQA (Quality)","CDDAMLRiskQAQualityClick")
		Call ClickLink("lnkAppealAssessment","Appeal Assessment","AppealAssessmentClick")
		Call clickWebElement("elmAMLAppealAssessment","Appeal Assessment","AMLAppealAssessmentClick")								'Introduced in sprint 40
		Call Clicklink("lnkRepairQualityCheckQA","RepairQualityCheckQA","RepairQualityCheckQAClick")
		Call ClickAllButton("pgAMLAssessment","Reviewed","ReviewedClick")
		Call selectRadioButton("rdAMLQCOutcome","AML QC Outcome","AMLQCOutcomeSelect")		
		Call selectRadioButton("rgpPCMPAppealAction","AML QC Outcome","AMLAppealActionSelect")	
		Wait 2
		Call selectRadioButton("rdAMLQCOutcomeSecond","AML QC Outcome Second","AMLQCOutcomeSecondSelect")
		Call clickWebElement("elmPCMPFailureReason","Failure Reason","AppealRequiredSelect")
		Wait 2
		Call selectRadioButton("rdAppealRequired","Appeal Required","AppealRequiredSelect")
		Call selectRadioButton("rgpAMLAppealAction","Appeal Action","AppealActionSelect")

		Call clicklink("lnkAddReason","Add Reason","AddReasonClick")
		Wait 2
		Call VerifyItems("lstReasonCode","Reason Code","ReasonCodeVerify")
		Call selectList("lstReasonCode","Reason Code","ReasonCodeSelect")
		wait 2
		Call VerifyItems("lstCategory","Category","CategoryVerify")
		Call selectList("lstCategory","Category","CategorySelect")
		'Call entertext("edtTimesFailed","Times Failed","TimesFailedSet")                   'Field Removed from application
		
		Call setcheckbox("chkFailHousekeeping","Fail Housekeeping","FailHousekeepingSet")
		Call setcheckbox("chkFailPolicy","Fail Policy","FailPolicySet")
		'Call setcheckbox("chkSeverity1","Severity 1","SeveritySet")
		Call selectList("lstPCMPSeverity","Severity","SeveritySet")

        Call Entertext("edtNotes","Notes","NotesSet")
		Call Entertext("edtRepairNote","RepairNote","RepairNoteSet")
		Call Entertext("edtEffortTime","EffortTime","EffortTimeSet")
		Call Clickbutton("btnSubmit","Submit","SubmitClick")
		Call verifyProperty("elmErrorType","Error Type","EXIST=[ErrorTypeVerify]")
   	End If

	If UCASE(getTestDataValue("AMLRiskAssessmentStatusCheck")) <> "SKIP"Then
			Call verifyProperty("elmAMLStatus","AML Status","EXIST=[AMLStatusVerify]")
	End If

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function

'========================================================================================='
'Product Fulfilment
'========================================================================================='
Public Function ProductFulfilment()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ProductFulfilment;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","CaseIDSearchCheck")
		Call verifypage("pgPegaCaseManagerPortal","Product Fulfilment","CaseIDSearchCheck")
		Call CaseID("SELECT")
		wait 5
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)		
		wait 10
		Call fireEvent("elmSearch","Search","SearchClick")
		'Call clickWebElement("elmSetupProducts","Setup Products","SetupProductsClick")
	End If
'	For i=0 TO 1
'		If i=0 Then
			Call clickWebElement("elmSetupProductsCA","Setup Products Current Account","Click")
'		Else
'			Call clickWebElement("elmSetupProductseBanking","Setup Products: ebanking","Click")
'		End If
			If UCASE(getTestDataValue("SetupProductCheck")) <> "SKIP"Then
				Call syncbrowser("brwPegaCaseManagerPortal","SetupProductCheck")
				Call verifypage("pgPegaCaseManagerPortal","Setup Product","SetupProductCheck")
				Call Entertext("edtNotes","Notes","Automation Testing Note")
				Call Entertext("edtEffortTime","EffortTime","1")
				Wait 3
				Call Clickbutton("btnSubmit","Submit","Click")
			End If
'	If UCASE(getTestDataValue("ApplicationScreenedCheck")) <> "SKIP"Then
'		Call syncbrowser("brwPegaCaseManagerPortal","ApplicationScreenedCheck")
'		Call verifypage("pgPegaCaseManagerPortal","Application Screened","ApplicationScreenedCheck")
'        Call Entertext("edtNotes","Notes","Automation Testing Note")
'		Call Entertext("edtEffortTime","EffortTime","1")
'		Wait 3
'		Call Clickbutton("btnSubmit","Submit","Click")
'   	End If
'	If UCASE(getTestDataValue("SetupInProgressCheck")) <> "SKIP"Then
'		Call syncbrowser("brwPegaCaseManagerPortal","SetupInProgressCheck")
'		Call verifypage("pgPegaCaseManagerPortal","Setup In Progress","SetupInProgressCheck")
'        Call Entertext("edtNotes","Notes","Automation Testing Note")
'		Call Entertext("edtEffortTime","EffortTime","1")
'		Wait 3
'		Call Clickbutton("btnSubmit","Submit","Click")
'   	End If
			If UCASE(getTestDataValue("SetupCompleteCheck")) <> "SKIP"Then
				Call syncbrowser("brwPegaCaseManagerPortal","SetupCompleteCheck")
				Call verifypage("pgPegaCaseManagerPortal","Setup Complete","SetupCompleteCheck")
'				If i=0  Then
					Call Entertext("edtSortCodeSetupComplate","Sor tCode SetupComplate","SortCodeSetupComplateSet")
					Call Entertext("edtAccountNumberSetupComplate","Account Number SetupComplate","AccountNumberSetupComplateSet")
					Call Entertext("edtDateOpenedSetupComplate","Date Opened SetupComplate","DateOpenedSetupComplateSet")
'					Call Entertext("edtDateOpenedSetupComplate","Date Opened SetupComplate","7/18/2016")
'				End If
				Call Entertext("edtNotes","Notes","Automation Testing Note")
				Call Entertext("edtEffortTime","EffortTime","1")
				Wait 3
				Call Clickbutton("btnSubmit","Submit","Click")
			End If
			wait 2
			Call Entertext("edtEffortTime","EffortTime","1")
			Wait 2
			Call Clickbutton("btnSubmit","Submit","Click")
			wait 2
'			If i =0 Then
				Call Clicklink("lnkAwaitingSMDUpload","Awaiting SMD Upload","AwaitingSMDUploadClick")
				If UCASE(getTestDataValue("AwaitingSMDUploadCheck")) <> "SKIP"Then
					Call syncbrowser("brwPegaCaseManagerPortal","AwaitingSMDUploadCheck")
					Call verifypage("pgPegaCaseManagerPortal","Awaiting SMD Upload","AwaitingSMDUploadCheck")
					Call Setcheckbox("chkMandateLoadedSMD","Mandate Loaded SMD","MandateLoadedSMDSet")
					Call Setcheckbox("chkAllDocumentsAttached","All Documents Attached","AllDocumentsAttachedSet")
					Call Entertext("edtNotes","Notes","Automation Testing Note")
					Call Entertext("edtEffortTime","EffortTime","1")
					Wait 3
					Call Clickbutton("btnSubmit","Submit","Click")
				End If
'			End If
			If UCASE(getTestDataValue("ClientReadinessCheck")) <> "SKIP"Then
				Call syncbrowser("brwPegaCaseManagerPortal","ClientReadinessCheck")
				Call verifypage("pgPegaCaseManagerPortal","Client Readiness","ClientReadinessCheck")
				Call Entertext("edtNotes","Notes","Automation Testing Note")
				Call Entertext("edtEffortTime","EffortTime","1")
				Wait 3
				Call Clickbutton("btnSubmit","Submit","Click")
			End If
			Call Clickbutton("btnGClose","Close","Click")
'	Next
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function

'========================================================================================='
'Resolve Case
'========================================================================================='
Public Function ResolveCase()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ResolveCase;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgAMLAssessment","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
		Call fireEvent("elmSearch","Search","SearchClick")	
		Call VerifyProperty("elmCaseStatus","Case Status","EXIST=[TRUE]")
		Call ClickLink("lnkResolveCase(Handover)","Resolve Case (Handover)","ResolveCaseHandoverClick")
	End If

	If UCASE(getTestDataValue("ResolveCaseCheck")) <> "SKIP"Then
		Call syncbrowser("brwPegaCaseManagerPortal","ResolveCaseCheck")
		Call verifypage("pgAMLAssessment","Resolve Case","ResolveCaseCheck")
		Call setCheckBox("chkKPIsUpdated","KPIs Updated","KPIsUpdatedSet")
		Call setCheckBox("chkConfirmAllDocuments","Confirm All Documents","ConfirmAllDocumentsSet")
		Call PegaFooter("PegaFooterCheck")
		Call Clickbutton("btnClose","Close Case","CloseCaseClick")
   	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function


'========================================================================================='
'Reports All
'========================================================================================='
Public Function ReportsAll()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ReportsAll;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("PegaCaseManagerCheck")) <> "SKIP"Then
		Call CaseID("SELECT")
        Call syncbrowser("brwPegaCaseManagerPortal","PegaCaseManagerCheck")
		Call verifypage("pgPegaCaseManagerPortal","Pega Case Manager","PegaCaseManagerCheck")
		Call ClickWebElement("elmReports","Reports","ReportsClick")
	End If

	If UCASE(getTestDataValue("ReportsAllCheck")) <> "SKIP"Then
		Call syncbrowser("brwReportsAll","ReportsAllCheck")
		Call verifypage("pgReportsAll","Reports All","ReportsAllCheck")
		Call Clicklink("linkOnboardingReports","Onboarding Reports","OnboardingReportsClick")
		Call Clicklink("lnkAMLAssessment","AML Assessment","AMLAssessmentClick")
		

		Setting.WebPackage("ReplayType") = 2 
		Wait 5
		Call Entertext("edtSearchReports","Search Reports","SearchReportsSet")
		Wait 2
		Setting.WebPackage("ReplayType") = 1
		
		Set wshobj =CreateObject("Wscript.Shell")
		wshobj.SendKeys"{TAB}"
		Set wshobj=Nothing

		Call Clicklink("lnkCasesDailyDashboard","Cases Daily Dashboard","CasesDailyDashboardClick")
		Call Clicklink("lnkInteractionDetails","Interaction Details","InteractionDetailsClick")
		Call Clicklink("lnkWithdrawnCases","Withdrawn Cases","WithdrawnCasesClick")
		Call Clicklink("lnkAMLCaseDetails","AML Case Details","AMLCaseDetailsClick")
		Call Clicklink("lnkAMLCasesSummaryByTask","AML Cases Summary By Task","AMLCasesSummaryByTaskClick")
		Call Clicklink("lnkAMLQACasesSummaryBy","AML QA Cases Summary By","AMLQACasesSummaryByClick")
		Call Clicklink("lnkAMLQCCompletedCases","AML QC Completed Cases","AMLQCCompletedCasesClick")
		Call Clicklink("lnkAMLRACompletedCases","AML RA Completed Cases","AMLRACompletedCasesClick")
		Call Clicklink("lnkRACompletedOnboarding","RA Completed Onboarding","RACompletedOnboardingClick")
'		Call Clicklink("lnkEffortPerAMLCaseDetailed","Effort Per AML Case Detailed","EffortPerAMLCaseDetailedClick")

		If UCASE(getTestDataValue("DisplayingRecordsVerify")) <> "SKIP"Then
				Call VerifyCaseInReport("wtbDisplayingRecords:lnkNext",getTestDataValue("DisplayingRecordsVerify"),gstrCaseID)
		End If

		Call VerifyAMLcaseSummary("wtbAMLCasesSummary","AML Cases Summary","AMLCasesSummaryVerify")
   	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function
'========================================================================================='
'Action For  RO Create Application
'========================================================================================='
Public Function CreateROApplication()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=CreateROApplication;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	Call VerifyPage("pgPegaROPortal","Get started with the application","PegaROPortalCheck")
	Call ClickButton("btnRequestClientApplication","Request Client Application","RequestClientApplicationClick")
	Wait 2
	Call EnterCompanyName("edtCompanyName","Company Name","ROCompanyNameSet")
	Call Entertext("edtCompanyName","Company Name","ManualROCompanyNameSet")
	Call EnterText("edtPrimarySortCode","Primary Sort Code","PrimarySortCodeSet")
    Call EnterText("edtOUID","OUID","OUIDSet")
'	Call VerifyItems("lstROArea","RO Area","ROAreaSelect")
	Call selectDropdown("lstROArea","RO Area","ROAreaSelect")

'	Call VerifyItems("lstCustomerTypeRO","Customer Type RO","CustomerTypeROSelect")
	Call selectDropdown("lstCustomerTypeRO","Customer Type RO","CustomerTypeROSelect")
	wait 4
	'Call VerifyItems("lstCustomerSubTypeRO","Customer Sub Type RO","CustomerSubTypeROSelect")
	Call selectDropdown("lstCustomerSubTypeRO","Customer Sub Type RO","CustomerSubTypeROSelect")
	Wait 2
	Call EnterText("edtRegisteredCharityNumberRO","Registered CharityNumber RO","RegisteredCharityNumberROSet")
'	Call VerifyItems("lstCharityRegisteredLocationRO","Charity Registered Location RO","CharityRegisteredLocationROSelect")
	Call selectDropdown("lstCharityRegisteredLocationRO","Charity Registered Location RO","CharityRegisteredLocationROSelect")

'	Call VerifyItems("lstEntityTypeRO","RO Entity Type","EntityTypeSelect")
'	Call SelectList("lstEntityTypeRO","RO Entity Type","EntityTypeSelect")
	Wait 2
'	Call VerifyItems("lstBrand","Brand","BrandSelect")
'	Call selectDropdown("lstBrand","Brand","BrandSelect")
	Call selectRadioButton("rdBusinessType","Business Type","Startup")
	wait 4
	Call SelectRadioButton("rdBrandRO","Brand","BrandSelect")
	wait 2
	Call VerifyItems("lstCustomerSegment","Customer Segment","CustomerSegmentSelect")
	Call selectDropdown("lstCustomerSegment","Customer Segment","CustomerSegmentSelect")
	wait 2
	Call VerifyItems("lstSchoolTypeRO","School Type","SchoolTypeSelect")
	Call selectDropdown("lstSchoolTypeRO","School Type","SchoolTypeSelect")

	Call selectRadioButton("rgpPCMPAccountOpenSpecificDate","AccountOpenSpecificDate","Yes")
	wait 4
	Call selectList("lstPCMPClientReqTargetDate","Client Require Date","Other")
	wait 4
	Call EnterPegaFormatDate("edtTargetDate","Target Date","TargetDateSet")
	wait 2
	Call EnterPegaFormatDate("edtClientMeetingDate","Client Meeting Date","ClientMeetingDateSet")
	wait 2
	Call EnterPegaFormatDate("edtClarificationDate","Clarification Date","ClarificationDateSet")
	wait 2
	Call selectRadioButton("rgpPCMPNeedLIP","Need LIP","No")

	Call ClickButton("btnCreateApplication","Create Application","CreateApplicationClick")
    wait 2
    Call GetCaseID("EleCaseID","Case ID","CaseIDGet")
	Call CaseID("UPDATE")
	Call VerifyProperty("elmApplicationCreated"," elmApplicationCreated","EXIST=[TRUE]")	
	Call ClickWebElement("EleCloseCase","Close Case ID","CloseCaseClick")

'	Call ClickLink("lnkLogout","Logout Application","LogoutClick")
	
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function


'========================================================================================='
'Action For  SUBMIT RO  Application
'========================================================================================='
Public Function SubmitROApplication()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=SubmitROApplication;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='

		If UCASE(getTestDataValue("PegaROPortalCheck")) <> "SKIP"Then
				Call VerifyPage("pgPegaROPortal"," RO Dashboard","PegaROPortalCheck")
				wait 4
				Call VerifyProperty("eleResearchStatus"," Case Status","EXIST=[ResearchStatusVerify]")				
				Call ClickCompanyName("lnkCompanyNameFirst","Company Name First","CompanyNameFirstClick")
		End If

		wait 5

		Call Clicklink("lnkApplicationConfirmationRO","Application Confirmation RO","Click")
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		'Business Details section
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("BusinessDetailsCheck")) <> "SKIP"Then
				Call VerifyProperty("eleBusinessDetails","Business Details Title","EXIST=[TRUE]")
				Call ClickWebElement("eleBusinessDetails","Business Details","BusinessDetailsClick")
				Call VerifyPage("pgPegaROPortal","Business Details","BusinessDetailsCheck")
				Call SelectRadioButton("rdBusinessType","Business Type","BusinessTypeUCASet")
				Wait 2
				Call Entertext("edtROSICCode","RO SIC Code","ROSICCodeSet")
				Call EnterText("edtGroupName","Group Name","GroupNameSet")
				Call SelectRadioButton("rdExistAccountHolder","Exist Account Holder","ExistAccountHolderSelect")
				wait 2
				Call ClickLink("lnkAddNatureOfBusinessRO","Add NatureOf BusinessRO","AddNatureOfBusinessROClick")
				wait 2
				Call EnterText("edtNatureOfBusinessRO","Nature Of Business RO","NatureOfBusinessROSet")
				Wait 4
				Call ClickLink("lnkAddReason","Add Reason Link","AddReasonClick")
				Wait 4
				Call SelectList("lstProductListRO","Product List RO","ProductListROSelect")
				Wait 4
				Call SelectList("lstReasonRO","Reason RO","ReasonROSelect")				
				Call SelectRadioButton("rdInterview","Interview Done","InterviewSelect")
				Call EnterPegaFormatDate("edtDateOfVisit","Date Of Visit","InterviewDateSet")
		End If

		'-------------------------------------------------------------------------------------------------------------------------------------------------
		'Expected Account Activity
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("ExpectedAccountActivityCheck")) <> "SKIP"Then
				Call VerifyProperty("elmExpectedAccountActivityTitle","Expected Account Activity Title","EXIST=[TRUE]")
				Call ClickWebElement("elmExpectedAccountActivityTitle","Expected Account Activity","ExpectedAccountActivityClick")
				Call VerifyPage("pgPegaROPortal","Expected Account Activity","ExpectedAccountActivityCheck")
				Call Entertext("edtAnnualTurnover","Annual Turnover","AnnualTurnoverSet")
		End If

		'-------------------------------------------------------------------------------------------------------------------------------------------------
		'Source Of Funds
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("SourceOfFundsCheck")) <> "SKIP"Then
				Call VerifyProperty("elmSourceOfFundsTitle","Source Of Funds Title","EXIST=[TRUE]")
				Call ClickWebElement("elmSourceOfFundsTitle","Source Of Funds","SourceOfFundsClick")
				Call VerifyPage("pgPegaROPortal","Source Of Funds","SourceOfFundsCheck")
				Call setCheckBox("chkSourcesOfFunds","chkSourcesOfFunds","SourcesOfFundsSet")
		End If

		''-------------------------------------------------------------------------------------------------------------------------------------------------
		'Individual Section
		''-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("IndividualsAssignPrimaryCheck")) <> "SKIP"Then
				Call VerifyProperty("eleIndividualsAssignPrimary","Individuals Assign Primary Title","EXIST=[TRUE]")
				Call ClickWebElement("eleIndividualsAssignPrimary","Individuals Assign primary","IndividualsAssignPrimaryClick")
				Call VerifyPage("pgPegaROPortal","Individuals Assign Primary","IndividualsAssignPrimaryCheck")
				wait 2
				Call IndividualDetailsRO("wtbIndividualsDetailsRO","Individual Details RO","IndividualsDetailsVerify")
		End If
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		'Products Section
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("ProductsCheck")) <> "SKIP"Then
				Call VerifyProperty("eleProducts","Products Title","EXIST=[TRUE]")
				Call ClickWebElement("eleProducts","Products","ProductsClick")
				Call VerifyPage("pgPegaROPortal","Products","ProductsCheck")
				Call ProductsSelectionRO("pgPegaCaseManagerPortal","Products Selecttion","ProductsSelecttionClick")      
		End If
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		' Authorized Representatives
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("AuthorisedRepresentativeCheck")) <> "SKIP"Then
	            Call VerifyProperty("eleAuthorisedRepresentative","Authorised Representative Title","EXIST=[TRUE]")
				Call ClickWebElement("eleAuthorisedRepresentative","Authorised Representative","AuthorisedRepresentativeClick")
				Call VerifyPage("pgPegaROPortal","Authorised Representative","AuthorisedRepresentativeCheck")
				wait 2
				Call SelectRadioButton("rdAR1","Authorized Represenatative One","AuthorizedRepresenatativeSelect")
				Call SelectRadioButton("rdAR2","Authorized Represenatative Two","AuthorizedRepresenatativeSelect")
				Call SelectRadioButton("rdAR3","Authorized Represenatative Three","AuthorizedRepresenatativeSelect")
		End If

		'-------------------------------------------------------------------------------------------------------------------------------------------------
		' Attached Documents
		'-------------------------------------------------------------------------------------------------------------------------------------------------
		If UCASE(getTestDataValue("AttachedDocumentsCheck")) <> "SKIP"Then
	            Call VerifyProperty("elmAttachedDocuments","Attached Documents Title","EXIST=[TRUE]")
				Call ClickWebElement("elmAttachedDocuments","Attached Documents","AttachedDocumentsClick")
				Call VerifyPage("pgPegaROPortal","Attached Documents","AttachedDocumentsCheck")
				Call SelectRadioButton("rdEducationDocumentsOne","Education Documents One","SMEEducationSelect")
				Call SelectRadioButton("rdEducationDocumentsTwo","Ownership Structure Attached","SMEEducationSelect")
				Call SelectRadioButton("rdEducationDocumentsThree","Members And Liabilities Letter","SMEEducationSelect")
				Call SelectRadioButton("rdEducationDocumentsFour","Ofsted Edubase and ISI Report","SMEEducationSelect")
		End If

		If UCASE(getTestDataValue("OtherActionsROCheck")) <> "SKIP"Then
				Call clickButton("btnOtherActionsRO","Other Actions RO","OtherActionsROClick")
				Call ClickWebElementFromWebTable("wtbOtherActionsRO","Other Action RO table","OtherActionsTableROClick")

				If UCASE(getTestDataValue("WithdrawROCheck")) <> "SKIP"Then
						Call VerifyPage("pgPegaROPortal","Withdraw Case RO","WithdrawCaseCheck")
						wait 5
						Call selectList("lstWithdrawReasonsRO","Withdraw Reason RO","WithdrawReasonROSelect")
				End If		
		End If

		Call Clickbutton("btnNextAdditionalInformation","Next Additional Information","NextAdditionalInformationClick")
		wait 10
		Call ClickButton("btnSubmitRO","Submit RO","SubmitROClick")
		Call VerifyProperty("elmApplicationWithdraw"," Application Withdraw","EXIST=[ApplicationWithdrawVerify]")
		Call ClickButton("btnSubmitToCWT","Submit To CWT","SubmitToCWTClick")
		wait 10
		Call clickWebElement("elmCloseMailCase","Close Mail Case","CloseMailCaseClick")
		Call VerifyProperty("eleConfirmationMsgRO"," Confirmation message of RO application submission","EXIST=[ConfirmationMsgROVerify]")
		
        wait 5
		
		Call ClickLink("lnkLogout","Logout Application","LogoutROClick")

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function

'========================================================================================='
'Action For  AttachFileRO
'========================================================================================='
Public Function AttachFileRO()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=AttachFileRO;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='	
		If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
			Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
			Call CaseID("SELECT")
			Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
			Call fireEvent("elmSearch","Search","Click")
			Call clicklink("lnkAddFiles","Add Files","Click")
		End If

		Call ClickLink("lnkCompanyNameFirst","Company Name Link","CompanyNameFirstClick")
		wait 5
		Call VerifyPage("pgPegaROPortal"," Attach File RO","AttachFileROCheck")

        Call ClickButton("btnAttachFile","Attach a File","AttachFileClick")
		Call EnterText("edtDescriptionAttachFile","Description AttachFile","DescriptionAttachFileSet")        
		Setting.WebPackage("ReplayType") = 2
		Call fireEvent("edtBrowserAttachFile","Browse AttachFile","BrowseAttachFileClick")
		Call EnterText("edtFileNameWin","File Name","FileNameWinSet")
		Call ClickButton("btnOpenWin","Open File","OpenWinClick")
		Call ClickButton("btnOKAttachFileRO","OK AttachFile RO","OKAttachFileROClick")
		Setting.WebPackage("ReplayType") = 1

		Call VerifyProperty("lnkCOB_DocumentsAttached"," COB_Documents Attached","EXIST=[COB_DocumentsAttachedVerify]")
		Call ClickWebElement("EleCloseCase","Close Case ","CloseCaseClick")
		wait 2
		Call ClickLink("lnkLogout","Logout Application","LogoutClick")
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()

End Function

'========================================================================================='
'Action For  Verify OfficeNotes
'========================================================================================='
Public Function VerifyDAndB()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=VerifyDAndB;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
				Call ClickLink("lnkCompanyNameFirstVerify","Company Name Link","CompanyNameFirstClick")
				wait 5
				
				' Business Details
				Call ClickWebElement("eleBusinessDetails","Business Details","BusinessDetailsClick")
				'company information
				Call VerifyObjectProperty("elmCompanyNameBusinessRO","Company Name Business Details Section","CompanyNameBusinessROVerify")
				Call VerifyObjectProperty("elmCompanyIDBusinessRO","Company ID Business Details Section","CompanyIDBusinessROVerify")
				'Trading address verification
				Call VerifyObjectProperty("elmAddressLine1","Line1 Trading Address Business Details Section","Line1TradingROVerify")
				Call VerifyObjectProperty("elmCity","City Trading Address  Business Details Section","CityTradingROVerify")
				Call VerifyObjectProperty("elmCounty","County Trading Address Business Details Section","CountyTradingROVerify")
				Call VerifyObjectProperty("elmPostCode","Post Code Trading Address  Business Details Section","PostCodeTradingROVerify")
				
				'  Individuals 
				Call ClickWebElement("eleIndividualsAssignPrimary","Individuals Assign primary","IndividualsAssignPrimaryClick")
				wait 2 
				Call VerifyAllTableDetailsItems("tblIndividual"," Individual RO Section","IndividualVerify")
				wait 2
				'office notes
				Call ClickWebElement("eleOfficeNotes","Office Notes","OfficeNotesClick")
				Call VerifyProperty("eleFinancialInformation","Financial Information","EXIST=[FinancialInformationVerify]")
				Call VerifyAllTableDetails("tblTurnoverStatement", "Turnover Statement", "TurnoverStatementVerify")
				Call VerifyAllTableDetails("tblSicCodes", "Sic Codes", "SicCodesVerify")
				
				Call ClickWebElement("EleCloseCase","Close Case ","CloseCaseClick")


'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()

End Function

'========================================================================================='
'Action For  Verify VerifyOtherActions
'========================================================================================='
Public Function OtherActions()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult
	bResult = True
	nIDIndex = 0
	Call getData("Table=OtherActions;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
			If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
				Call verifypage("pgPegaCaseManagerPortal","Pega Case Manager Portal","CaseIDSearchCheck")
				Call CaseID("SELECT")
	
				Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)	
				
				Call fireEvent("elmSearch","Search","SearchClick")
				wait 5
				Call clickButton("btnActions","Actions","ActionsClick")
		
				Call clickWebElementFromWebTable("tblAction","Action table","TableActionClick")
				wait 2
			End If

			Call Clickbutton("btnUpdateClient'sApplication","Update Client's Application","OAUpdateClientsApplicationClick")

			'suspend case
			If UCASE(getTestDataValue("SuspendCaseCheck")) <> "SKIP"Then
				    Call ClickElement("eleSuspendWork","Suspend Work","Suspend WorkClick")
					Call verifypage("pgOtherActions","Suspend Case","SuspendCaseCheck")
					Call selectRadioButton("rdPendTypeCWT","Pend Type CWT","PendTypeCWTSelect")
					Call selectList("lstPendReasonCWT","Pend Reason CWT","PendReasonCWTSelect")
					str=dateadd("d",1, Date) &" " &left(time,5)
					Call entertext("edtSuspendUntil","Suspend Until",str)
					Call setCheckBox("chkPendOnlyCWT","Pend Only CWT","PendOnlyCWTSet")
			End If
			'Add a party
			If UCASE(getTestDataValue("AddPartyCheck")) <> "SKIP"Then
					Call verifypage("pgOtherActions","Add Party","AddPartyCheck")
					Call selectList("lstParty","Party","PartySelect")
					Call clickButton("btnAddParty","Add Party","AddPartyClick")
					wait 2
					Call clickButton("eleThirdElement","Open Party","ThirdElementClick")
					Call entertext("edtWorkParty","Work Party","WorkPartySet")
					wait 2
			End If
		'Record Interactions
			If UCASE(getTestDataValue("RecordInteractionsCheck")) <> "SKIP"Then
					Call verifypage("pgOtherActions","Record Interactions","RecordInteractionsCheck")
					Call ClickLink("lnkAddInteraction","Add Interaction","AddInteractionClick")
					Call entertext("edtPartyInteraction","Party Interaction","PartyInteractionSet")
					Call selectRadioButton("rdPartyInteraction","Party Interaction","PartyInteractionSelect")
					Call entertext("edtFullName","Full Name","FullNameSet")
					Call entertext("edtCollegueContacted","Collegue Contacted","CollegueContactedSet")
					wait 2
			End If
		' Withdraw
			If UCASE(getTestDataValue("WithdrawCheck")) <> "SKIP"Then
						Call verifypage("pgOtherActions","Withdraw","WithdrawCheck")
						Call selectList("lstWithdrawReasonCWT","Withdraw Reason","WithdrawReasonCWTSelect")
			End If

		 ' Nominate Primary Contact
			If UCASE(getTestDataValue("NominatePrimaryContactCheck")) <> "SKIP"Then
						Call verifypage("pgOtherActions","Nominate Primary Contact","NominatePrimaryContactCheck")
						Call selectRadioButton("rdNominatePrimaryContact"," Nominate Primary Contact","NominatePrimaryContactSet")
			End If

			 ' Capture the CRA rating
			If UCASE(getTestDataValue("CaptureCRARatingCheck")) <> "SKIP"Then
						Call verifypage("pgOtherActions","Capture the CRA rating","CaptureCRARatingCheck")
						Call entertext("edtCRARating","CRA Rating","CRARatingSet")
			End If

				 ' Transfer Assignment
			If UCASE(getTestDataValue("TransferAssignmentCheck")) <> "SKIP"Then
						Call verifypage("pgOtherActions","Transfer Assignment","TransferAssignmentCheck")
						Call Selectlist("lstReassignOperator","Reassign Operator","ReassignOperatorSelect")
			End If

				 ' Send Client Invite 
			If UCASE(getTestDataValue("SendClientInviteCheck")) <> "SKIP"Then
						Call verifypage("pgOtherActions","Transfer Assignment","SendClientInviteCheck")
'						Call verifyProperty("elmSendClientInviteTitle","Send ClientInvite Title","EXIST=[TRUE]")
'						Call verifyProperty("elmClientInviteKAPDetails","elmClientInviteKAPDetails","EXIST=[TRUE]")
						Call setAllCheckBox("pgOtherActions","KAP","ON")
			End If

			Call Entertext("edtEffortTime","EffortTime","EffortTimeSet")
			Call clickButton("btnSubmit","Submit","SubmitClick")
			Call verifyProperty("elmTransferActionPerformed","Transfer Action Performed","EXIST=[TransferActionPerformedVerify]")
			Call clickButton("btnClose","Close","CloseClick")
			If Browser("brwOtherActions").Page("pgOtherActions").Frame("Page---->OtherActions").WebButton("btnDiscardOA").Exist(2) Then
				Browser("brwOtherActions").Page("pgOtherActions").Frame("Page---->OtherActions").WebButton("btnDiscardOA").Click
			End If

			Call ClickWebElement("elmCloseMailCase","Close Case ","CloseCaseClick")

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()

End Function

'========================================================================================='
'Action For  Login Clicent
'========================================================================================='
Public Function LoginClient()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=LoginClient;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("LoginClientCheck")) <> "SKIP"Then
		Call verifypage("pgLoginClient","Login Client","LoginClientCheck")
		Call Entertext("edtUsernameClient","User Name","UserNameSet")
		Call Entertext("edtPasswordClient","Password","PasswordSet")
		Call Clickbutton("btnLogInClient","Log In","LogInClick")

'		Call Entertext("edtPassword1","Password 1","Password1Set")
'		Call Entertext("edtPassword2","Password 2","Password1Set")
'		Call Clickbutton("btnSubmitNewPassword","Submit New Password","SubmitNewPasswordClick")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function

'========================================================================================='
'Action For  Login Clicent
'========================================================================================='
Public Function ClientDashboard()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ClientDashboard;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='

   	If UCASE(getTestDataValue("ChooseApplicationCheck")) <> "SKIP"Then
		Call verifypage("pgChooseApplication","Choose Application","ChooseApplicationCheck")
'		Call ClickLink("lnkComplanyNameClient","Complany Name Client","ComplanyNameClientClick")
	End If

	If UCASE(getTestDataValue("ContactPreferenceCheck")) <> "SKIP"Then
		Call verifypage("pgContactPreference","Contact Preferences","ContactPreferenceCheck")
		Call setAllCheckBox("pgContactPreference","Contact Preferences","ContactPreferencesSet")
		Call Clickbutton("btnContinueCP","Continue","ContinueCPClick")
	End If

	If UCASE(getTestDataValue("AssignRolesAndPermissionsCheck")) <> "SKIP"Then
		Call verifypage("pgAssignRolesAndPermissions","Assign Roles And Permissions","AssignRolesAndPermissionsCheck")
		Wait 5
		Call setAllCheckBox("pgAssignRolesAndPermissions","Assign Roles And Permissions","AssignRolesAndPermissionsSet")
		Call Clickbutton("btnContinueRol","Continue","ContinueRoleClick")
	End If

	If UCASE(getTestDataValue("ApproveMyAgreementCheck")) <> "SKIP"Then
		Call verifypage("pgApproveMyAgreement","Approve My Agreement","ApproveMyAgreementCheck")
        Call Clickbutton("btnApproveMyAgreement","Approve My Agreement","ApproveMyAgreementClick")
	End If

	If UCASE(getTestDataValue("ReviewCheck")) <> "SKIP"Then
		Call verifypage("pgReview","Review","ReviewCheck")
        Call Clickbutton("btnConfirmDetails","Confirm Details","ConfirmDetailsClick")
	End If

	If UCASE(getTestDataValue("DataUsageCheck")) <> "SKIP"Then
		Call verifypage("pgDataUsage","Data Usage","DataUsageCheck")
		Call setCheckBox("chkAgreementDataUsage","Agreement Data Usage","AgreementDataUsageSet")
        Call Clickbutton("btnAcceptDataUsagePolicy","Accept Data Usage Policy","AcceptDataUsagePolicyClick")
	End If

	If UCASE(getTestDataValue("ApprovalCheck")) <> "SKIP"Then
		Call verifypage("pgApproval","Approval","ApprovalCheck")
		Call setCheckBox("chkAgreementDataApproval","Agreement Data Approval","AgreementDataApprovalSet")
        Call Clickbutton("btnApproveAgreement","Approve Agreement","ApproveAgreementClick")
	End If

	If UCASE(getTestDataValue("ConfirmationCheck")) <> "SKIP"Then
		Call verifypage("pgConfirmation","Confirmation","ConfirmationCheck")
		Call setCheckBox("chkAgreementDataConfirm","Agreement Data Confirmation","AgreementDataConfirmSet")
        Call Clickbutton("btnApproveAgreementConfim","Approve Agreement Confim","ApproveAgreementConfimClick")
	End If

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()

End Function
'========================================================================================='
'Action For  Login Clicent Mandate Verification
'========================================================================================='
Public Function ClientMandateVerification()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=ClientDashboard;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='

   	If UCASE(getTestDataValue("ChooseApplicationCheck")) <> "SKIP"Then
		Call verifypage("pgChooseApplication","Choose Application","ChooseApplicationCheck")
'		Call ClickLink("lnkComplanyNameClient","Complany Name Client","ComplanyNameClientClick")
	End If

	If UCASE(getTestDataValue("ContactPreferenceCheck")) <> "SKIP"Then
		Call verifypage("pgContactPreference","Contact Preferences","ContactPreferenceCheck")
		Call setAllCheckBox("pgContactPreference","Contact Preferences","ContactPreferencesSet")
		Call Clickbutton("btnContinueCP","Continue","ContinueCPClick")
	End If

	If UCASE(getTestDataValue("AssignRolesAndPermissionsCheck")) <> "SKIP"Then
		Call verifypage("pgAssignRolesAndPermissions","Assign Roles And Permissions","AssignRolesAndPermissionsCheck")
		Wait 5
		Call setAllCheckBox("pgAssignRolesAndPermissions","Assign Roles And Permissions","AssignRolesAndPermissionsSet")
		Call Clickbutton("btnContinueRol","Continue","ContinueRoleClick")
	End If

	If UCASE(getTestDataValue("ApproveMyAgreementCheck")) <> "SKIP"Then
		Call verifypage("pgApproveMyAgreement","Approve My Agreement","ApproveMyAgreementCheck")
        Call Clickbutton("btnApproveMyAgreement","Approve My Agreement","ApproveMyAgreementClick")
	End If

	If UCASE(getTestDataValue("ReviewCheck")) <> "SKIP"Then
		Call verifypage("pgReview","Review","ReviewCheck")
        Call Clickbutton("btnConfirmDetails","Confirm Details","ConfirmDetailsClick")
	End If

	If UCASE(getTestDataValue("DataUsageCheck")) <> "SKIP"Then
		Call verifypage("pgDataUsage","Data Usage","DataUsageCheck")
		Call setCheckBox("chkAgreementDataUsage","Agreement Data Usage","AgreementDataUsageSet")
        Call Clickbutton("btnAcceptDataUsagePolicy","Accept Data Usage Policy","AcceptDataUsagePolicyClick")
	End If

	If UCASE(getTestDataValue("ApprovalCheck")) <> "SKIP"Then
		Call verifypage("pgApproval","Approval","ApprovalCheck")
		Call setCheckBox("chkAgreementDataApproval","Agreement Data Approval","AgreementDataApprovalSet")
        Call Clickbutton("btnApproveAgreement","Approve Agreement","ApproveAgreementClick")
	End If

	If UCASE(getTestDataValue("ConfirmationCheck")) <> "SKIP"Then
		Call verifypage("pgConfirmation","Confirmation","ConfirmationCheck")
		Call setCheckBox("chkAgreementDataConfirm","Agreement Data Confirmation","AgreementDataConfirmSet")
        Call Clickbutton("btnApproveAgreementConfim","Approve Agreement Confim","ApproveAgreementConfimClick")
	End If

'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()

End Function

'========================================================================================='
'Action For  PegaHeader
'========================================================================================='
Public Function PegaHeader()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=PegaHeader;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("HeaderCheck")) <> "SKIP"Then
		Call verifypage("pgHeader","Header","HeaderCheck")
		Call VerifySibling("elmCompany","Company","CompanyVerify")
		Call VerifySibling("elmCompanyID","Company ID","CompanyIDVerify")
		Call VerifySibling("elmRelationshipOwner","Relationship Owner","RelationshipOwnerVerify")
		Call VerifySibling("elmOnboardingManager","Onboarding Manager","OnboardingManagerVerify")
		Call VerifySibling("elmInitialClientMeeting","InitialClientMeeting","InitialClientMeetingVerify")
		Call VerifySibling("elmTotalEffort","Total Effort","TotalEffortVerify")
		Call VerifySibling("elmROConfirmation","RO Confirmation","ROConfirmationVerify")
		Call VerifySibling("elmROConfirmationDate","RO Confirmation Date","ROConfirmationDateVerify")
		
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function



'========================================================================================='
'Action For  PegaAuditLogs
'========================================================================================='
Public Function PegaAuditLogs()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
	Call getData("Table=PegaAuditLogs;Columns=*")
	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
	If UCASE(getTestDataValue("CaseIDSearchCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Case ID Search","CaseIDSearchCheck")
		Call CaseID("SELECT")
		Call EnterText("edtSearchWorkflowID","Search Workflow ID",gstrCaseID)
		Call fireEvent("elmSearch","Search","SearchClick")
		Call ClickWebelement("elmAuditLogs","Audit Logs","AuditLogsClick")
	End If

	If UCASE(getTestDataValue("AuditLogsCheck")) <> "SKIP"Then
		Call verifypage("pgPegaCaseManagerPortal","Audit Logs","AuditLogsCheck")
		Call VerifyAuditLogs("wtbAuditLogTable:lnkAuditLogNext","Audit Log","AuditLogVerify")
	End If
'======================================================================================='    
	If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
	UnloadOR()
End Function


'========================================================================================='
'Action For  Extrenal QA Manager
'========================================================================================='
Public Function ExtrenalQAManager()
	On Error Resume Next
	If objErr.number = 11 Then
		Exit Function
	End If
	Dim bResult, nTemp,strAppURL,strApplicationName
	bResult = True
	nIDIndex = 0
'	loadOR "PegaE2E.tsr"
	While nIDIndex <= Ubound(arrData)	
'======================================================================================='
		Set fso=CreateObject("Scripting.FileSystemObject")
		If Right( Environment.Value("OS"),1)<>"7" Then
			folderPath="C:\Documents and Settings\"&Environment.Value("UserName")&"\My Documents\Downloads\"
			csvDocPath=folderPath&"ExportDataCSV.csv"
			If fso.FolderExists(folderPath) Then
					fso.DeleteFile(folderPath&"*"),True
			Else 
					fso.CreateFolder(folderPath)
			End If
			excelDocPath=folderPath&"ExportData.xls"
			csvDocPath=folderPath&"ExportDataCSV.csv"
		Else
			folderPath="C:\Users\"&Environment.Value("UserName")&"\Downloads\"
			If fso.FolderExists(folderPath) Then
					fso.DeleteFile(folderPath&"*"),True
			Else 
					fso.CreateFolder(folderPath)
			End If
			excelDocPath=folderPath&"ExportData.xls"
			csvDocPath=folderPath&"ExportDataCSV.csv"
		End If

		Call ClickWebelement("elmPCMPReports","Reports","Click")
		Call ClickLink("lnkPCMPQASample","QA Sample","Click")
		Call ClickLink("lnkQSQCCompleted"," QC Completed","Click")
		Call EnterText("edtQSUpdateData","Update Data","Today")
		Call Clickbutton("btnQSApplyChanges","Apply Changes","Click")
		wait 1
		Call ClickLink("lnkQSExportToExcel"," Export To Excel","Click")
		wait 7
		Set WshShell = CreateObject("WScript.Shell")
		WshShell.SendKeys "{ENTER}" 
		wait 4
'		Set objExcel = CreateObject("Excel.Application")
'		excelDocPath="C:\Documents and Settings\"&Environment.Value("UserName")&"\My Documents\Downloads\ExportData.xls"
'		csvDocPath="C:\Documents and Settings\"&Environment.Value("UserName")&"\My Documents\Downloads\ExportDataCSV.csv"
'		Set objWorkbook = objExcel.Workbooks.Open(excelDocPath)
'		objExcel.DisplayAlerts = FALSE
'		objExcel.Visible = FALSE
'
'		i=4
'		Do Until objExcel.Cells(i, 1).Value = gstrCaseID
'			Set objRange = objExcel.Cells(i, 1).EntireRow
'		   objRange.Delete
'		Loop
'		objExcel.ActiveWorkbook.Save	
'		set st=objWorkbook.Worksheets("ExportData")
'        st.SaveAs csvDocPath, xlCSV
'		'objWorkbook.SaveAs csvDocPath ,fileformat=xlCSV
'		objExcel.Quit
'		Set objExcel=Nothing

		Const xlCSV = 6
		Set objExcel = CreateObject("Excel.Application")
'		Set fso=CreateObject("Scripting.FileSystemObject")
'		If Right( Environment.Value("OS"),1)<>"7" Then
'			folderPath="C:\Documents and Settings\"&Environment.Value("UserName")&"\My Documents\Downloads\"&"UploadDoc\"
'			If fso.FolderExists(folderPath) Then
'					fso.DeleteFile(folderPath&"*"),True
'			Else 
'					fso.CreateFolder(folderPath)
'			End If
'			excelDocPath=folderPath&"ExportData.xls"
'			csvDocPath=folderPath&"ExportDataCSV.csv"
'		Else
'			folderPath="C:\Users\"&Environment.Value("UserName")&"\Downloads\"&"UploadDoc\"
'			If fso.FolderExists(folderPath) Then
'					fso.DeleteFile(folderPath&"*"),True
'			Else 
'					fso.CreateFolder(folderPath)
'			End If
'			excelDocPath=folderPath&"ExportData.xls"
'			csvDocPath=folderPath&"ExportDataCSV.csv"
'		End If
		Set objWorkbook = objExcel.Workbooks.Open(excelDocPath)
		objExcel.DisplayAlerts = FALSE
		objExcel.Visible = FALSE
	
		i=4
		Do Until objExcel.Cells(i, 1).Value = gstrCaseID
			Set objRange = objExcel.Cells(i, 1).EntireRow
		   objRange.Delete
		Loop
			i=1
		Do Until i=3
			Set objRange = objExcel.Cells(1, 1).EntireRow
		   objRange.Delete
		   i=i+1
		Loop

		'objExcel.ActiveWorkbook.Save
		Set objWorksheet = objWorkbook.Worksheets("ExportData")
		objWorksheet.SaveAs csvDocPath , xlCSV
		objExcel.Quit
		wait 2
		Call closeBrowser("brwQASample"," QA Sample","Close")
		wait 1
		Call ClickWebelement("elmReportsClose","Reports Close","Click")
		wait 1
		Call ClickLink("lnkPCMPUploadQACases"," Upload QA Cases","Click")
		wait 3
		Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").WebFile("ebfUploadFile").Set csvDocPath
		wait 1
		Call Clickbutton("btnPCMPUpload","Upload","Click")
		wait 1
		Call Clickbutton("btnClose"," Close","Click")
		wait 30
		Call ClickLink("lnkPCMPBU"," BU","Click")
		For i =0 to 30
			If Browser("brwPegaCaseManagerPortal").Page("pgPegaCaseManagerPortal").Link("lnkPCMPRefresh").Exist(2) Then
				Call Clickbutton("lnkPCMPRefresh"," Refresh","Click")
				wait 2
			Else
				Exit For
			End If
		Next
		wait 2
		Call Clickbutton("btnSubmit"," Submit","Click")
		wait 1
		Call Clickbutton("btnClose"," Close","Click")



'======================================================================================='
If objErr.number = 11 Then
			nIDIndex =  Ubound(arrData) + 1
		Else
			nIDIndex = nIDIndex + 1
	End If	
	Wend
'	UnloadOR()
End Function 


