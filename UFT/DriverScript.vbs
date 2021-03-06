'General Header
'#######################################################################################################################
'Script Description        : Driver class for the framework
'Test Tool/Version        : HP Quick Test Professional 10+
'Test Tool Settings        : N.A.
'Application Automated    : Flight Application
'Author                    : Cognizant
'Date Created            : 21/11/2012
'#######################################################################################################################
Option Explicit    'Forcing Variable declarations

Dim gobjDriverScript: Set gobjDriverScript = New DriverScript

'#######################################################################################################################
'Class Description       : Driver class which encapsulates the core logic of the CRAFT framework
'Author                    : Cognizant
'Date Created            : 09/11/2012
'#######################################################################################################################
    Class DriverScript
    
    Private m_dtmStartTime, m_dtmEndTime
    Private m_intCurrentIteration, m_intCurrentSubIteration
    Private m_arrBusinessFlowData()
    
    
    '###################################################################################################################
    'Function Description   : Function to drive the test execution
    'Input Parameters         : None
    'Return Value            : None
    'Author                    : Cognizant
    'Date Created            : 11/10/2012
    '###################################################################################################################
       Public Sub DriveTestExecution()
       	Environment("strCurrentKeyword") = ""
       	Environment("FieldID") = 1
        Startup()
        'If Instr(gobjTestParameters.CurrentTestcase,"CLA") <> 0 Then
        	gobjTestParameters.CurrentTestcase = GetTestCaseNames()
        'End If
        arrTcs = Split(gobjTestParameters.CurrentTestcase,",")
        For inctcs = 0 To Ubound(arrTcs)   'Removing the empty Tcs name due to ,
        	gobjTestParameters.CurrentTestcase = arrTcs(inctcs)
        	InitializeTestIterations()
	        InitializeTestReport()
	        InitializeDataTable()
	'        RegisterUserDef()
	        InitializeBusinessFlow()
	        Set objTCRun = gobjQCUtility.Create_TestRuns	        
	        ExecuteTestIterations()
	        gobjQCUtility.UploadTestResults(objTCRun)
	'        UnRegisterUserDef()
	        'CloseBrowsers()
	        'QC_Add_TestRuns()
'	        intTestCaseID = gobjDataTable.GetData("Business_Flow","TestCaseID")'Environment.Value("TestID")
'   			intTestSetID =  gobjDataTable.GetData("Business_Flow","TestSetID") 'Environment.Value("TestSetId")
	        WrapUp()
        Next
        ExitRun
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
       Private Sub Startup()
        m_dtmStartTime = Now()
        SetDefaultTestParameters()
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub SetDefaultTestParameters()
        gobjTestParameters.CurrentTestcase = Environment.Value("TestName")
        
'        If Instr(1,Environment.Value("TestName"),"E2E") = 0  Then
            gobjTestParameters.CurrentScenario = gobjQCUtility.GetCurrentTestParentFolder()
'        Else
        	
'            gobjTestParameters.CurrentScenario = GetScenario
'        End If
        
        gobjTestParameters.IterationMode = TestArgs("IterationMode")
        gobjTestParameters.StartIteration = TestArgs("StartIteration")
        gobjTestParameters.EndIteration = TestArgs("EndIteration")
    End Sub
    '###################################################################################################################
    Private Function GetCurrentApplication()        
        Dim strFilePath, strConnectionString, r_objConn,Cmd, strQuery
        strFilePath = "C:\Users\COG12193\Desktop\CompassAutomationSuite\EndToEnd\Run Manager.xls"
        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
        strQuery = "Select Execute_App from [Functional$] where TestCase='" & Environment.Value("TestName") & "'"
'        strQuery = "Select * from [Functional$]"
        Set Cmd = CreateObject("ADODB.Command")
        Cmd.ActiveConnection = r_objConn
        Cmd.CommandText = strQuery
        Dim objTestData: Set objTestData = CreateObject("ADODB.Recordset")
        objTestData.CursorLocation = 3
        objTestData.Open Cmd                         
        Set objTestData.ActiveConnection = Nothing
        r_objConn.Close
        GetCurrentApplication = objTestData.Fields(0)        
    End Function        

	'#####################################################################################################################
	Private Function GetScenario()
		Dim strScenarioFolder, arrSplitPath, Scenario,crntApplication
'		strScenarioFolder = Environment.Value("TestDir")
'		arrSplitPath = Split(strScenarioFolder,"\")
		Scenario = gobjQCUtility.GetCurrentTestParentFolder()
		crntApplication = GetCurrentApplication()
		GetScenario = crntApplication & "_" & Scenario
	End Function
	'#####################################################################################################################
    
    '###################################################################################################################
    Private Sub InitializeTestIterations()
        Select Case gobjTestParameters.IterationMode
            Case "RunOneIterationOnly"
                gobjTestParameters.StartIteration = 1
                gobjTestParameters.EndIteration = 1
            Case "RunRangeOfIterations"
                If (gobjTestParameters.StartIteration) > (gobjTestParameters.EndIteration) Then
                    Err.Raise 6002, "CRAFT", "StartIteration cannot be greater than EndIteration"
                End If
                If (gobjTestParameters.StartIteration = "") Then
                    gobjTestParameters.StartIteration = 1
                End If
                If (gobjTestParameters.EndIteration = "") Then
                    gobjTestParameters.EndIteration = 1
                End If
            Case "RunAllIterations"
                gobjExcelDataAccess.DatabasePath = GetDatatablesPath(gobjTestParameters.CurrentScenario)
                gobjExcelDataAccess.DatabaseName = gobjTestParameters.CurrentScenario
                gobjExcelDataAccess.Connect()
                
                Dim strCurrentTestCase, strTestDataSheet, strQuery, objTestData
                strCurrentTestCase = gobjTestParameters.CurrentTestcase
                strTestDataSheet = Environment.Value("DefaultDataSheet")
                Set objTestData = CreateObject("ADODB.Recordset")
                strQuery = "Select Distinct Iteration from [" & strTestDataSheet & "$]" &_
                                                    " where TC_ID='" & strCurrentTestCase & "'"
                Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
                gobjExcelDataAccess.Disconnect()
                
                Dim intIterationCount
                intIterationCount = objTestData.RecordCount
                If intIterationCount = 0 Then
                    Err.Raise 6003, "CRAFT", "The specified test case " & strCurrentTestCase &_
                                                    " is not found in the default test data sheet!"
                End If
                
                'Release all objects
                objTestData.Close
                Set objTestData = Nothing
                
                gobjTestParameters.StartIteration = 1
                gobjTestParameters.EndIteration = intIterationCount
        End Select
        
        m_intCurrentIteration = gobjTestParameters.StartIteration
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Function GetDatatablesPath(strDataTableName)
        Dim strDataFilePath, intPathLength
        strDataFilePath = PathFinder.Locate("[ALM\Resources] Resources\Compass_Automation_Suite\Datatables\" & strDataTableName & ".xls") 
        
        'strDataFilePath = PathFinder.Locate("C:\Users\COG12193\Desktop\Testedata\Datatables\" & strDataTableName & ".xls") 
        intPathLength = Len(strDataFilePath) - Len("\" & strDataTableName & ".xls")
        GetDatatablesPath = Left(strDataFilePath, intPathLength)
    End Function
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub InitializeTestReport()
        InitializeReportSettings()
        
        gobjReport.InitializeReport()
        gobjReport.InitializeTestLog()
        gobjReport.AddTestLogHeading(gobjReportSettings.ProjectName & " - " &_
                                        gobjReportSettings.ReportName & " - Automation Execution Results")
        gobjReport.AddTestLogSubHeading "Date & Time",  ": " & Now(), _
                                        "Iteration Mode", ": " & gobjTestParameters.IterationMode
        gobjReport.AddTestLogSubHeading "Start Iteration", ": " & gobjTestParameters.StartIteration, _
                                        "End Iteration",  ": " & gobjTestParameters.EndIteration
        gobjReport.AddTestLogTableHeadings()
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub InitializeReportSettings()
        gobjReportSettings.ReportPath = SetUpTempResultFolder()
        gobjReportSettings.ReportName = gobjTestParameters.CurrentScenario & "_" & gobjTestParameters.CurrentTestcase
        gobjReportSettings.ProjectName = Environment.Value("ProjectName")
        gobjReportSettings.LogLevel = Environment.Value("LogLevel")
        gobjReportSettings.ExcelReport = Environment.Value("ExcelReport")
        gobjReportSettings.HtmlReport = Environment.Value("HtmlReport")
        gobjReportSettings.TakeScreenshotPassedStep = Environment.Value("TakeScreenshotPassedStep")
        gobjReportSettings.TakeScreenshotFailedStep = Environment.Value("TakeScreenshotFailedStep")
        gobjReportSettings.LinkScreenshotsToTestLog = False
        gobjReportSettings.ReportTheme = Environment.Value("ReportsTheme")
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Function SetUpTempResultFolder()
        Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
        
        Dim strTempResultPath    'Using the Windows temp folder to store the results before uploading to QC
        strTempResultPath = objFso.GetSpecialFolder(2) & "\Run_mm-dd-yyyy_hh-mm-ss_XX"
        
        'Create Temp results folder if it does not exist
        If Not objFso.FolderExists (strTempResultPath) Then
            objFso.CreateFolder(strTempResultPath)
        End If
        
        strTempResultPath = strTempResultPath & "\" & gobjTestParameters.CurrentTestcase
        
        'Delete test case level result folder if it already exists
        If objFso.FolderExists(strTempResultPath) Then
            objFso.DeleteFolder(strTempResultPath)
            
            'Wait until the folder is successfully deleted
            Do While(1)
                If Not objFso.FolderExists(strTempResultPath) Then
                    Exit Do
                End If
            Loop
        End If
        
        'Create separate folder with the test case name
        objFso.CreateFolder(strTempResultPath)
        
        SetUpTempResultFolder = strTempResultPath
        
        'Release all objects
        Set objFso = Nothing
    End Function
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub InitializeDataTable()
        gobjDataTable.DataTablePath = GetDatatablesPath(gobjTestParameters.CurrentScenario)
        gobjDataTable.CommonDataTablePath = GetDatatablesPath("Common Testdata")
        gobjDataTable.DataTableName = gobjTestParameters.CurrentScenario
        gobjDataTable.DataReferenceIdentifier = Environment.Value("DataReferenceIdentifier")
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub InitializeBusinessFlow()
        Dim strBusinessFlowSheet, strCurrentTestCase
        strBusinessFlowSheet = "Business_Flow"
        strCurrentTestCase = gobjTestParameters.CurrentTestcase
        
        gobjExcelDataAccess.DatabasePath = gobjDataTable.DataTablePath
        gobjExcelDataAccess.DatabaseName = gobjDataTable.DataTableName
        gobjExcelDataAccess.Connect()
        
        Dim strQuery, objTestData
        Set objTestData = CreateObject("ADODB.Recordset")
        objTestData.CursorLocation = 3
        strQuery = "Select * from [" & strBusinessFlowSheet & "$] where TC_ID='" & strCurrentTestCase & "'"
        Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
        gobjExcelDataAccess.Disconnect()
        
        If objTestData.RecordCount = 0 Then
            Err.Raise 6004, "CRAFT", "Testcase '" & strCurrentTestCase & "' not found in the 'Business_Flow' sheet!"
        End If
        
        ReDim m_arrBusinessFlowData(500)    ' Maximum size of a record fetched from Excel 
        Dim intColumnCount, intArrVal, strCrntScenario, strScenariocnt
        Dim arrScenario()
        intArrVal = 0
        For intColumnCount = 1 to (objTestData.Fields.Count - 1)
'            If IsNull(objTestData(intColumnCount).Value) Or Trim(objTestData(intColumnCount).Value) = "" Then
'                'ReDim Preserve m_arrBusinessFlowData(intColumnCount - 2)
'                Exit For
'            End If
			If Instr(1,objTestData(intColumnCount).Name,"TestCaseID")<> 0 Then			
					Environment("TestCaseID") = objTestData(intColumnCount).Value			
			Elseif Instr(1,objTestData(intColumnCount).Name,"TestSetID")<> 0  Then
					Environment("TestSetID") = objTestData(intColumnCount).Value
			Elseif Instr(1,objTestData(intColumnCount).Name,"Dry_Run")<> 0  Then
					Environment("Dry_Run") = objTestData(intColumnCount).Value
			End if
			
			
            If Instr(1,objTestData(intColumnCount).Name,"Keyword") <> 0 and Trim(objTestData(intColumnCount).Value) <> "" Then
            	 If IsNull(objTestData(intColumnCount).Value) Or Trim(objTestData(intColumnCount).Value) = "" Then
	                'ReDim Preserve m_arrBusinessFlowData(intColumnCount - 2)
	                Exit For
	            End If	
				strScenario = Split (objTestData(intColumnCount).Value , "_")(0) 
					if strScenario = "Sce" then
						arrSceiteration = Split (objTestData(intColumnCount).Value , ",")
						If Ubound(arrSceiteration) > 0 Then
							intSceiteration = arrSceiteration(1)
						Else  
							intSceiteration = 1
						End If						
						strCrntScenario = Split (objTestData(intColumnCount).Value , ",")(0)
						gobjExcelDataAccess.DatabasePath = gobjDataTable.DataTablePath
					    gobjExcelDataAccess.DatabaseName = gobjDataTable.DataTableName
				       	gobjExcelDataAccess.Connect()					        
				        Dim strSceQuery, objScenario, intSceColumnCount
				        Set objScenario = CreateObject("ADODB.Recordset")
				        objScenario.CursorLocation = 3
				        strSceQuery = "Select * from [Scenario$] where Scenario='" &  strCrntScenario & "'"
				        Set objScenario = gobjExcelDataAccess.ExecuteQuery(strSceQuery)
				        gobjExcelDataAccess.Disconnect()
				        
				        If objScenario.RecordCount = 0 Then
				            Err.Raise 6004, "CRAFT", "Scenario '" & strCrntScenario & "' not found in the 'Scenario' sheet!"
				  		End If
					  	For intSceinc = 1 To intSceiteration
						  	For intSceColumnCount = 1 to (objScenario.Fields.Count-1) ' Removed first column not to include Scenario name 
								If Trim(objScenario(intSceColumnCount).Value) = "" or isnull(objScenario(intSceColumnCount).Value) Then
									Exit for
								End If
								m_arrBusinessFlowData(intArrVal) = objScenario(intSceColumnCount).Value
				     			intArrVal = intArrVal + 1 
				     			'ReDim Preserve m_arrBusinessFlowData(ubound(m_arrBusinessFlowData())+1)
				     		Next
						Next			     		
				Else			
	            	 m_arrBusinessFlowData(intArrVal) = objTestData(intColumnCount).Value
				     intArrVal = intArrVal + 1 
					'ReDim Preserve m_arrBusinessFlowData(ubound(m_arrBusinessFlowData())+1)			     
				End If
            End If
              
        Next
        
        'Release all objects
        objTestData.Close
        Set objTestData = Nothing
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub ExecuteTestIterations()
        Do While (m_intCurrentIteration <= gobjTestParameters.EndIteration)
            gobjReport.AddTestLogSection("Iteration: " & m_intCurrentIteration)
            
            If Instr(Environment.Value("ResultDir"), "TempResults") = 0_
            And Environment.Value("OnError") <> "NextStep" Then
                On Error Resume Next
            End If
            
            ExecuteTestCase()
            
            ExceptionHandler()
            
            m_intCurrentIteration = m_intCurrentIteration + 1
        Loop
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub ExecuteTestCase()
        If Ubound(m_arrBusinessFlowData) < 0 Then
            Err.Raise 6005, "CRAFT", "The business flow for the testcase '" & strCurrentTestCase & "' is empty!"
        End If
        
        Dim objKeywordDirectory : Set objKeywordDirectory = CreateObject("Scripting.Dictionary")
        
        Dim intCurrentKeywordNum, intKeywordIterations, intCurrentKeywordIteration
        Dim arrCurrentFlowData, strCurrentKeyword
        
        For intCurrentKeywordNum = 0 to UBound(m_arrBusinessFlowData)
        	If m_arrBusinessFlowData(intCurrentKeywordNum) <> "" Then
	            arrCurrentFlowData = Split(m_arrBusinessFlowData(intCurrentKeywordNum), ",")
	            strCurrentKeyword = arrCurrentFlowData(0)
	    
	            If UBound(arrCurrentFlowData) = 0 Then
	                intKeywordIterations = 1
	             Else
	                intKeywordIterations = arrCurrentFlowData(1)
	            End If
	            
	            For intCurrentKeywordIteration = 0 to (intKeywordIterations - 1)
	                If objKeywordDirectory.Exists(strCurrentKeyword) Then
	                    objKeywordDirectory.Item(strCurrentKeyword) = objKeywordDirectory.Item(strCurrentKeyword) + 1
	                Else
	                    objKeywordDirectory.Add strCurrentKeyword, 1
	                End If
	                m_intCurrentSubIteration = objKeywordDirectory.Item(strCurrentKeyword)        
	                Environment("Subiteration_Current") = m_intCurrentSubIteration
	                gobjDatatable.SetCurrentRow gobjTestParameters.CurrentTestCase,_
	                                            m_intCurrentIteration,_
	                                            m_intCurrentSubIteration
	                
	                Dim strSectionDescription
	                If (m_intCurrentSubIteration > 1) Then
	                    gobjReport.AddTestLogSubSection strCurrentKeyword &_
	                                                        " (SubIteration : " & m_intCurrentSubIteration & ")"
	                Else
	                    gobjReport.AddTestLogSubSection strCurrentKeyword
	                End If
	                
	                'InvokeBusinessComponent strCurrentKeyword
	                Environment("strCurrentKeyword") = strCurrentKeyword
	                Call Object_Flow(strCurrentKeyword)
	                 'Call PaymentDetails()
	'                strQuery = "Select * from [Scenario$] where Scenario = '" & strCurrentKeyword & "'"
	'                Set RstKeywords = ExtractSceKeywords(strQuery)
	'                ExecuteKeywords(RstKeywords)
	            Next
	           End If
	        Next
       
        objKeywordDirectory.RemoveAll()
        Set objKeywordDirectory = Nothing
    End Sub    
    '##################################################################################################################
        Function ExtractSceKeywords (strQuery)
        'CheckPreRequisites()
        gobjExcelDataAccess.Connect()
        Set objTestData = CreateObject("ADODB.Recordset")
        objTestData.CursorLocation = 3
        Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
        gobjExcelDataAccess.Disconnect()
'        Set adodb = CreateObject("ADODB.Connection")
'        DataFile_Path = GetDatatablesPath(gobjTestParameters.CurrentScenario)
'        DataFile_Path = DataFile_Path & "\NBV_SmokeTestingSuite.xls"
'        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataFile_Path &_
'                                                                    ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
''        adodb.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
''        "Data Source=" & DadaFile_Path & ";" & _
''        "Extended Properties=""Excel 8.0;HDR=Yes"";" 
'        Set objSceKeywords = CreateObject("ADODB.Recordset")
'        'strQuery = "Select * from [Scenario$] where Scenario = '" & strCurrentKeyword & "'"
'        Set objSceKeywords = adodb.Execute(strQuery)
'        Set ExtractSceKeywords= objSceKeywords
'        Set adodb = nothing
'        Set objSceKeywords = nothing
        Set ExtractSceKeywords = objTestData
    End Function
    '###################################################################################################################
    Private Sub InvokeBusinessComponent(strCurrentKeyword)    
        Execute strCurrentKeyword
    End Sub
    '###################################################################################################################
    Private Sub ExecuteKeywords(RstKeywords)
            Strfields = RstKeywords.Fields.count
            strCurrentKeyword = RstKeywords.Fields(0)
            
            For incKey = 1 To (Strfields-1)
                Execution_Flow = RstKeywords.Fields(inckey)
                If Execution_Flow <> "" Then
                    Call ExecuteFlow(strCurrentKeyword,Execution_Flow) 
                End If
                
                
            Next
    End Sub
    '###################################################################################################################
    Private Sub ExceptionHandler()
        If (Err.Number <> 0) Then
            'Error Reporting
            gobjReport.UpdateTestLog "Error", Err.Description, "Fail"
            
            'Error Response
            If TestArgs("StopExecution") Then
                gobjReport.UpdateTestLog "CRAFT Info", _
                                            "Test execution terminated by user! All subsequent tests aborted...", "Done"
                
                CustomErrorResponse()
                m_intCurrentIteration = gobjTestParameters.EndIteration
            Else
                Select Case Environment.Value("OnError")
                    Case "NextStep"
                        gobjReport.UpdateTestLog "CRAFT Info", _
                                                    "Refer QTP Results for full details regarding the error...", "Warning"
                        Err.Clear
                    Case "NextIteration"
                        gobjReport.UpdateTestLog "CRAFT Info", _
                                                    "Test case iteration terminated by user! " &_
                                                    "Proceeding to next iteration (if applicable)...", "Done"
                        CustomErrorResponse()
                    Case "NextTestCase"
                        gobjReport.UpdateTestLog "CRAFT Info", _
                                                    "Test case terminated by user! " &_
                                                    "Proceeding to next test case (if applicable)...", "Done"
                        
                        CustomErrorResponse()
                        m_intCurrentIteration = gobjTestParameters.EndIteration
                    Case "Stop"
                        TestArgs("StopExecution") = True
                        gobjReport.UpdateTestLog "CRAFT Info", _
                                                    "Test execution terminated by user! " &_
                                                    "All subsequent tests aborted...", "Done"
                        
                        CustomErrorResponse()
                        m_intCurrentIteration = gobjTestParameters.EndIteration
                End Select
            End If
        End If
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub CustomErrorResponse()
        'CloseFlightApp()
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub WrapUp()
        m_dtmEndTime = Now()
        CloseTestReport()
        Environment("strCurrentKeyword") = ""
        m_intCurrentSubIteration = "1"
        'UploadResultsToQc        
        'ExitRun
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub CloseTestReport()
       Dim strExecutionTime
       strExecutionTime = gobjUtil.GetTimeDifference(m_dtmStartTime, m_dtmEndTime)
       gobjReport.AddTestLogFooter strExecutionTime
    End Sub
    '###################################################################################################################
    
    
    '###################################################################################################################
    Private Sub UploadResultsToQc()
        If (gobjReportSettings.ExcelReport) Then
            gobjQCUtility.AttachFileToTestRun gobjReportSettings.ReportPath & "\Excel Results\" &_
                                                                            gobjReportSettings.ReportName & ".xls"
        End If
        
        If (gobjReportSettings.HtmlReport) Then
            gobjQCUtility.AttachFileToTestRun gobjReportSettings.ReportPath & "\HTML Results\" &_
                                                                            gobjReportSettings.ReportName & ".html"
        End If
        gobjQCUtility.AttachFolderToTestRun gobjReportSettings.ReportPath & "\Screenshots"
        'gobjQCUtility.QC_Add_TestRuns()
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Sub UploadDataTablesToQc()
        If Environment.Value("IncludeTestDataInReport") Then
            gobjQCUtility.AttachFileToTestRun PathFinder.Locate("Datatables\" & gobjDataTable.DataTableName & ".xls")
            gobjQCUtility.AttachFileToTestRun PathFinder.Locate("Datatables\Common Testdata.xls")
        Else
            Dim strTestResourcesPath
            strTestResourcesPath = GetTestResourcesFrameworkPath()
            Set gobjQCUtility.QcConnection = QCUtil.QCConnection
            gobjQCUtility.UpdateFileInTestResources gobjDataTable.DataTableName & ".xls",_
                                                        strTestResourcesPath & "\Datatables"
            gobjQCUtility.UpdateFileInTestResources "Common Testdata.xls",_
                                                        strTestResourcesPath & "\Datatables"
        End If
    End Sub
    '###################################################################################################################
    
    '###################################################################################################################
    Private Function GetTestResourcesFrameworkPath()
        Dim objQtpApp: Set objQtpApp = CreateObject("QuickTest.Application")
        GetTestResourcesFrameworkPath = objQtpApp.Folders.Item(1)
        Set objQtpApp = Nothing
    End Function
    '###################################################################################################################
	Private Function GetTestCaseNames()
		 		gobjExcelDataAccess.DatabasePath = GetDatatablesPath(gobjTestParameters.CurrentScenario)
               	gobjExcelDataAccess.DatabaseName = gobjTestParameters.CurrentScenario
                gobjExcelDataAccess.Connect()
                
                Dim strCurrentTestCase, strTestDataSheet, strQuery, objTestData
                'strCurrentTestCase = gobjTestParameters.CurrentTestcase
                strTestDataSheet = Environment.Value("DefaultDataSheet")
                Set objTestData = CreateObject("ADODB.Recordset")
                strQuery = "Select TC_ID from [General_Data$]" &_
                                                    " where Exec_Flag='Yes'"
                Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
                gobjExcelDataAccess.Disconnect()
                For incTcs = 1 To objTestData.RecordCount
                	strTcs = Strtcs & objTestData.Fields(0) & ","
                	objTestData.MoveNext
                Next
                GetTestCaseNames = Left(strTcs,len(strTcs)-1)
	End Function 
End Class

