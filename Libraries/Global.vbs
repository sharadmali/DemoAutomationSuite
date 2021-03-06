'====================================================================================================
'This VBS file contains all the global variables
'====================================================================================================
Public gstrApplicationURL
Public gstrCaseID
Public gstrUserNamePEGA


public gstrEnvironmentName	' Store Environment name
public grsObjectRepository	' Store each objects path
public gstrProject		' Project Namepublic gstrTester		' Tester Name
public gstrBaseDir		' Test Directory
Public gstrTestname		' Test name to execute
public gstrMachine		' Machine Name
public gstrLogFileName		' Logfile
public gstrErrFileName      	' Error File
public gstrAppURL     		' Application / URL
public gstrControlFileName    	' ControlFileName
public gstrCurScenario      	' Current Scenario
public gstrcurScenarioDesc    	' Current Scenario Description
public gstrCurBatch     	' Current Test batch Path
public gstrCurTSR     		' Current XML file path
public garrTestCaseNames    	' Test Case names for individual scenario
public gstrTestCaseName     	' Test Case name
public gstrTestCaseNameFromQC 	' Test Case name from QC
public gstrTCStartTime      	' Test Case execution start time
public gstrTCEndTime      	' Test Case execution end time
'public gobjObjectClass      	' To get object description from XML
public gstrTestDBName       	' Test Database name
public gstrEnvironment      	' For Test Environment
public gstrRunName      	' To store Run name
public gnCurTestDataID      	' The ID of the current test data item
public gstrResFilePath      	' Store the Result File Path
public gstrReportFilePath   	' Store the Report File Path
public ExecutionTime      	' Store the execution time of the script
public strObjName     		' Store the information of Object Name
public strKeyword     		' Store the information of Keyword
public strParam       		' Store the information of  Parameter
public gstrDescription      	' Store the Description
public gstrHref       		' Store the path of file containing screenshot
public gstrRunStartTime     	' Store start time of current Run
public gstrRunEndTime     	' Store end time of current Run
public gstrRunDate      	' Store execution date of current Run
public gintTotalScenarios   	' Store the no. of total scenarios executed
public gintTotalpass      	' Store the no. of Pass scenarios
public gintTotalFail      	' Store the no. of Fail scenarios
public gstrExecutionStatus    	' Store execution status of currently executed test case
public gstrDesc               	' Store the information about reporting statement
Public gScreenShotPath
Public gIterCombinedStatus 
Public gdetailReport
Public gObjectpath
Public gTempTCRowsCnt           ' setCurrentRow in ExcelSheet For Execution
Public gSummaryreport   	' view the summary report
Public gBforeandAftdit    	' path of the screen shot
Public gNameforBforandAftedit   ' name for screen shots before and after edit
Public gRowKey 
Public gStrTestStep
Public gstrQCDesc
Public gstrExpectedResult
Public gDataCount
Public gnRollNumber
Public gstrActionName
Public gstrActionDataSet
Public gExistCount
Public gbIterationFlag
Public gbOfferStatusFlag
Public rCurTime
Public gstrCurSummaryPath
Public gstrProjectName
Public gstrExecuteMethod
Public gstrProjectUser
Public gbReporting
Public dictApplicationURL
Public gbScreenShotImageStatus
'====================================================================================================

Public gstrGroupName
Public gstrEnvironmentSetupDir
Public gstrLibrariesDir
Public gstrObjectRepositoryDir
Public gstrRuntimeReportsDir
Public gstrScreenshotsDir
Public gstrSummaryReportDir
Public gstrTestResultLogDir
Public gstrActionFilesDir
Public gstrBatchFilesDir
Public gstrControlFilesDir
Public gstrTestDataDir
Public gTotalCount,gTotalPassed,gTotalFailed,gGroupExecutionTime
Public gstrAppName
Public nIDIndex
Public gDataID

'Application Specific Variables 
Public gstrPropVal
Public gstrAccessCode    	'Store OTP 
Public objDBConn			'To store the database connection
Public rsDBRecords			'To store the records returned by SQL
Public gstrDB2User 			'Store the UserID for DB2 Connection
Public gstrDB2Password		'Store the Password for DB2 Connection
Public gstrUserName			'Stores the UserID for the TC
Public gstrSC1				'Stores the SortCode for the Customer
Public gstrSC2				'Stores the SortCode for the Customer
Public gstrSC3				'Stores the SortCode for the Customer
Public gstrAccNo			'Stores the Account Number for the Customer
Public gstrSTCCurrentAcnt	'Stores the Current Account Number for the Customer for Save The Change
Public gstrSTCSavingsAcnt	'Stores the Savings Account Number for the Customer or Save The Change
Public gDictIBCData			'Stores the IBC details for the environment
Public gstrOTPSerialNumber	'Stores the OTP serial no
Public gstrHsdlURL		'Stores the HSDL URL
Public gstrDate			'Storeslateral Date for Payment Execution
Public strAccountName 
Public objAccount
Public objDormant
Public gbListValue
Public prevID
Public gstrUserID
Public gstrAccountBalance
Public gstrIBCSwitch
Public gbAppRefNum   ' Stores Value of Application reference Number for Creditcard
Public gstrApplicationName
Public gfrmDate ' stores payment date on confirmation page with shorter format.
Public gstrReferenceNumber
Public gstrMortgageRefNumber
Public gstrPartyID
Public gstrCardNumber
Public objOffers
Public gintInBasketCount		'Stores the case count of the Inbasket
Public gintCaseID		'gets the case ID
Public gobjHeader
Public gstrUser			'Stores the user name
Public gstrpropertyvalue  	'Stores the property value
Public QCConnection
Public treeMgr
Public gstrStrUsername
Public strtempusname
Public strtempusname1
Public gstrtemppswd
'Public QCSuitePath
'Public SuiteName
Public gstrStockCode
Public gstrNeonUserName
Public gstrNeonPassword
Public gDictUser
Public gDictAdam
Public gstrAdamCredentials
Public gstrAdamServer
Public gstrAdamUsername
Public gstrAdamPassword
Public gWaitCount
Public strApplicationName
Public gstrAccountCode
Public gstrOneStage
Public gDictEnvironment
Public  blnQCUpdation
Public gstrCurID
Public gstrQCIntegrationArray()
Public gstrEnv
Public gQCIntegration_Flag

Public gstrScreenshotDocPath 'Path of Screenshot Document
Public gstrDocName	      'Scrrenshot document name
Public gstrFoldersExistFlag   'Back up folder exist flag
public gstrScreenShotPath	'Path of Screenshot
Public gstrScreenshotBookmark 'Path to Screenshot document and bookmarkname
Public gstrBackupReportFlag  'Create Backup of Report Flag
Public gstrMoveHtmlReportFlag
Public gstrCreateScreenshotDocPathFlag
Public gstrVariableDefine    'Variable defination flag
Public gstrScreenshotBookmarkNum  'Bookmark number for Scrrenshot document
Public gstrScreenshotEND_OF_DOC  
Public gobjWord
Public gobjDoc
Public gobjSelect
Public gstrReportStatement
Public gResultLogtExeTime
Public gstrTotalExecutionTime
Public gstrDataReusabilityStatus
Public gstrDataTestCaseID
Public gstrDataTestCaseIteration
Public gstrBrowserURL
Public gstrBrowserHWND
Public gstrFileData
Public gstrREMAIL
Public gstrFirstName
Public gNotificationNumber
Public gActivationVerificationCode
Public gClientID
Public gClientName
Public gstrRegistrationEmail
Public gUserName
Public gstrPaymentID
Public gstrTemplateName
Public gstrJobID
Public gstrGlobalPasscode
Public gstrPassCode
Public gstrRunTimeCounter
Public gstrChallengeCode
Public gstrResponseCode
Public gstrSTPRegion
'Public gstrLogFileName
'Public gstrSTPRegion
Public gstrPuttyIPAddress
Public gstrPuttyIDPass
Public gstrPuttyInstance
Public gstrPuttyTitle
Public gstriiliadTRNNumber
Public gstrIFR_TRAN_REF_NO
Public gstrIFP_INSTRUCTION_ID

Public gstrOriginatorAccountNo
Public gstrDestinationAccountNo
  Public gstrDestinationSortCode
Public gstrAmount
Public gstrMappedHyperFileName
Public gstrTnum
Public gstrBankPaymentID
Public dictPuttyEnv
Public gstrBankHoliday

Public gstrCompany
Public gstrIndv
Public gstrCompanyName
Public gstrCustomerType