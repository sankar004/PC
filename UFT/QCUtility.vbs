'General Header
'#######################################################################################################################
'Script Description		: QCUtil class to manage integration with HP Quality Center/Application Lifecycle Management
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated	: N.A.
'Author					: Cognizant
'Date Created			: 04/07/2011
'#######################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gobjQCUtility: Set gobjQCUtility = New QCUtility

'#######################################################################################################################
'Class Description   	: Class to interact with QC/ALM
'Author					: Cognizant
'Date Created			: 23/09/2012
'#######################################################################################################################
Class QCUtility
	
	Private m_objQcConnection
	
	'###################################################################################################################
	Public Property Set QcConnection(objQcConnection)
		Set m_objQcConnection = objQcConnection
	End Property
	'###################################################################################################################
	
	
	'###################################################################################################################
	'Function Description   : Function to get the parent folder of the currently executing test
	'Input Parameters 		: None
	'Return Value    		: Parent folder of the currently executing test
	'Author 				: Cognizant
	'Date Created			: 09/10/2012
	'###################################################################################################################
	Public Function GetCurrentTestParentFolder()
		Dim objCurrentTest: Set objCurrentTest = QCUtil.CurrentTest
		GetCurrentTestParentFolder = objCurrentTest.Field("TS_SUBJECT").Name
		Set objCurrentTest = Nothing
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to get the grandparent folder of the currently executing test
	'Input Parameters 		: None
	'Return Value    		: Grandparent folder of the currently executing test
	'Author 				: Cognizant
	'Date Created			: 09/10/2012
	'###################################################################################################################
	Public Function GetCurrentTestGrandParentFolder()
		Dim objCurrentTest: Set objCurrentTest = QCUtil.CurrentTest
		strCurrentTestPath = objCurrentTest.Field("TS_SUBJECT").Path
		Dim arrSplitPath : arrSplitPath = Split(strCurrentTestPath,"\")
		GetCurrentTestGrandParentFolder = arrSplitPath(UBound(arrSplitPath)-1)
		Set objCurrentTest = Nothing
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to get the description of the currently executing test
	'Input Parameters 		: None
	'Return Value    		: Description of the currently executing test
	'Author 				: Cognizant
	'Date Created			: 09/10/2008
	'###################################################################################################################
	Public Function GetCurrentTestDescription()
		Dim objCurrentTest: Set objCurrentTest = QCUtil.CurrentTest
		GetCurrentTestDescription = objCurrentTest.Field("TS_DESCRIPTION")
		Set objCurrentTest = Nothing
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to get the specified user field value of the currently executing test
	'Input Parameters 		: None
	'Return Value    		: User field value of the currently executing test
	'Author 				: Cognizant
	'Date Created			: 17/07/2013
	'###################################################################################################################
	Public Function GetCurrentTestUserFieldValue(strUserField)
		Dim objCurrentTest: Set objCurrentTest = QCUtil.CurrentTest
		GetCurrentTestUserFieldValue = objCurrentTest.Field(strUserField)
		Set objCurrentTest = Nothing
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to update an existing file within the test resources module
	'Input Parameters		: strFileName, strResourceFolderPath
	'Return Value    		: None
	'Author					: Cognizant
	'Date Created			: 11/10/2012
	'###################################################################################################################
	Public Sub UpdateFileInTestResources(strFileName, strResourceFolderPath)
		Dim objQcResourceFolder: Set objQcResourceFolder = GetResourceFolderByPath(strResourceFolderPath)
		
		Dim objQcConnection: Set objQcConnection = m_objQcConnection
		Dim objQcResourceFactory: Set objQcResourceFactory = objQcConnection.QCResourceFactory
		Dim objQcResourceFilter: Set objQcResourceFilter = objQcResourceFactory.Filter
		objQcResourceFilter.Filter("RSC_PARENT_ID") = objQcResourceFolder.id
		objQcResourceFilter.Filter("RSC_NAME") = "'" & strFileName & "'"
		
		Dim strResourceFolderClientPath, strResourceFileClientPath
		strResourceFileClientPath = PathFinder.Locate(strResourceFolderPath & "\" & strFileName)
		strResourceFolderClientPath = Left(strResourceFileClientPath, Len(strResourceFileClientPath) - Len(strFileName) - 1)
		
		Dim objQcResourceList, objQcResource
		Set objQcResourceList = objQcResourceFilter.NewList()
		If objQcResourceList.Count = 1 Then
			Set objQcResource = objQcResourceList.Item(1)
			objQcResource.Filename = strFileName
			objQcResource.Post
			objQcResource.UploadResource strResourceFolderClientPath, True
		Else
			Err.Raise 5003, "QCUtility", "The given resource was not found in the test resources module!"
		End If
		
		'Release all objects
		Set objQcResource = Nothing
		Set objQcResourceList = Nothing
		Set objQcResourceFilter = Nothing
		Set objQcResourceFactory = Nothing
		Set objQcResourceFolder = Nothing
		Set objQcConnection = Nothing
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	Private Function GetResourceFolderByPath(strFolderPath)
		CheckQcConnection()
		
		Dim objQcConnection: Set objQcConnection = m_objQcConnection
		Dim objQcResourceFolderFactory: Set objQcResourceFolderFactory = objQcConnection.QCResourceFolderFactory
		Dim objQcResourceFolder: Set objQcResourceFolder = objQcResourceFolderFactory.Root
		
		'Navigate the resources tree to locate the datatable resource
		Dim intCurrentFolder, intCurrentResourceFolder
		intCurrentResourceFolder = 0
		Dim arrFolders
		arrFolders = Split(strFolderPath, "\")
		For intCurrentFolder = 0 To UBound(arrFolders)
			If Len(arrFolders(intCurrentFolder)) > 0 Then 'Skip over empty strings caused by leading/trailing "\"s as well as multiple "\"s
				For intCurrentResourceFolder = 1 To objQcResourceFolder.Count 'Iterate over the children of the current folder
					If objQcResourceFolder.Child(intCurrentResourceFolder).Type = 10 _
					And objQcResourceFolder.Child(intCurrentResourceFolder).Name = arrFolders(intCurrentFolder) Then
						Set objQcResourceFolder = objQcResourceFolder.Child(intCurrentResourceFolder)
						Exit For
					End If
				Next
			End If
		Next
		
		Set GetResourceFolderByPath = objQcResourceFolder
		
		If objQcResourceFolder.Name = objQcResourceFolderFactory.Root.Name Then
			Err.Raise 5002, "QCUtility", "The given folder was not found in the test resources module!"
		End If
		
		'Release all objects
		Set objQcResourceFolder = Nothing
		Set objQcResourceFolderFactory = Nothing
		Set objQcConnection = Nothing
	End Function
	'###################################################################################################################
	
	'###################################################################################################################
	Private Sub CheckQcConnection()
		If IsEmpty(m_objQcConnection) Then
			Err.Raise 5001, "QCUtility", "QC connection unavailable!"
		End If
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to get the list of files and folders within the specified Test Resources folder
	'Input Parameters		: strFolderPath
	'Return Value    		: None
	'Author					: Cognizant
	'Date Created			: 11/10/2012
	'###################################################################################################################
	Public Sub GetChildrenOfTestResourcesFolder(strFolderPath, arrChildFolderList(), arrChildFileList())
		Dim objQcResourceParentFolder: Set objQcResourceParentFolder = GetResourceFolderByPath(strFolderPath)
		Dim intFolderCount, intFileCount
		intFolderCount = 0
		intFileCount = 0
		
		Dim i, objQcResourceFolder, objQcResourceFile
		For i = 1 To objQcResourceParentFolder.Count	'Iterate over the children of the current folder
			If objQcResourceParentFolder.Child(i).Type = 10 Then
				Set objQcResourceFolder = objQcResourceParentFolder.Child(i)
				Redim Preserve arrChildFolderList(intFolderCount)
				arrChildFolderList(intFolderCount) = objQcResourceFolder.Name
				intFolderCount = intFolderCount + 1
			Else
				Set objQcResourceFile = objQcResourceParentFolder.Child(i)
				Redim Preserve arrChildFileList(intFileCount)
				arrChildFileList(intFileCount) = objQcResourceFile.FileName
				intFileCount = intFileCount + 1
			End If
		Next
		
		'Release all objects
		Set objQcResourceFolder = Nothing
		Set objQcResourceFile = Nothing
		Set objQcResourceParentFolder = Nothing
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to attach all files within the given folder to the current test run (in test lab)
	'Input Parameters		: strFolderPath
	'Return Value    		: None
	'Author					: Cognizant
	'Date Created			: 11/10/2012
	'###################################################################################################################
	Public Sub AttachFolderToTestRun(strFolderPath)
		Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
		Dim objFolder: Set objFolder = objFso.GetFolder(strFolderPath)
		Dim objFileList: Set objFileList = objFolder.Files
		Dim objFile
		For each objFile in objFileList
			AttachFileToTestRun objFile.Path
		Next
		
		'Release all objects
		Set objFile = Nothing
		Set objFileList = Nothing
		Set objFolder = Nothing
		Set objFso = Nothing
	End Sub
	'###################################################################################################################
	
	'###################################################################################################################
	'Function Description   : Function to attach the specified file to the current test run (in test lab)
	'Input Parameters		: strFilePath
	'Return Value    		: None
	'Author					: Cognizant
	'Date Created			: 11/10/2012
	'###################################################################################################################
	Public Sub AttachFileToTestRun(strFilePath)
		Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
		If Not objFso.FileExists(strFilePath) Then
			Err.Raise 5004, "QCUtility", "The given file to be attached is not found!"
		End If
		Set objFso = Nothing
		
		Dim objFoldAttachments: Set objFoldAttachments =  QCUtil.CurrentRun.Attachments
		Dim objFoldAttachment: Set objFoldAttachment = objFoldAttachments.AddItem(Null)
		objFoldAttachment.FileName = strFilePath
		objFoldAttachment.Type = 1
		objFoldAttachment.Post
		
		Set objFoldAttachment = Nothing
		Set objFoldAttachments = Nothing
	End Sub
	'###################################################################################################################
	
Function Create_TestRuns()
			Dim intTestCaseID, intTestSetID
    intTestCaseID = Environment("TestCaseID")'Environment.Value("TestID")
    intTestSetID =  Environment("TestSetID")'Environment.Value("TestSetId")

    'strAttachmentFilePath = "H:\run_results.html"
    blnAddToTSet = True
'    Select Case g_Flag
'    Case 0
'                strStatus="Passed"
'    Case Else
'                strStatus="Failed"
'    End Select
    If Isnumeric(intTestCaseID)=False and Isnumeric(intTestSetID) = False Then 'Checks whether Test case Id and Test Set ID is provided in the datafile
                    Reporter.ReportEvent micfail, "Uploading test run to QC failed. Either Test case Id or Test Set ID is not available in the datafile"
                    'Craft Results
                    Exit Function
    End if
   'Set up required QC connections
    Set objQCCon = QCUtil.QCConnection
    Set objTestFact = objQCCon.TestFactory
    Set objTestSetFact = objQCCon.TestSetFactory
    Set objTestitem = objTestFact.item(intTestCaseID)
    Set objTestSetitem = objTestSetFact.item(intTestSetID)
    Set objTestcase = objTestSetitem.TSTestFactory 
    Set objTestPlanFilter=objTestFact.Filter
    objTestPlanFilter("TS_TEST_ID")=intTestCaseID
    Set objTestList=objTestFact.NewList(objTestPlanFilter.Text)
    If objTestList.Count=0 Then
                Reporter.ReportEvent micfail, "Uploading test run to QC failed. Given test case ID not available in test plan"
                'Craft Results
                Exit Function
    End If
    Set objTstSetLst = objTestcase.NewList("") 'Checks whether the given test case is available in the test set
    For Each objTCItem In objTstSetLst
                If Int(objTCItem.field("TC_TEST_ID")) = Int(intTestCaseID)Then
                            blnAddToTSet = "False"
                            Set objTC=objTestcase.newlist(objTCItem.id) 
                            For Each objActTC in objTC
                                        Set objTCtoTS = objActTC
                            Next
                            Exit for
                End If
    Next
    If blnAddToTSet = True Then
                Set objTCtoTS=objTestcase.additem(null)
                objTCtoTS.Field("TC_TEST_ID")=objTestitem
                objTCtoTS.post
    End If
    Set objRunFact=objTCtoTS.RunFactory
    Set objTCRun=objRunFact.additem(objRunFact.UniqueRunName)
    objTCRun.Field("RN_DRAFT") = Environment("Dry_Run")
    objTCRun.Field("RN_RUN_NAME") = gobjTestParameters.CurrentTestcase
    objTCRun.Status="Not Completed"
    objTCRun.Post
'    objTCRun.CopyDesignSteps
'    objTCRun.Post
    Set Create_TestRuns = objTCRun
End Function
Function UploadTestResults(objTCRun)
	Select Case g_Flag
    Case 0
                strStatus="Passed"
    Case Else
                strStatus="Failed"
    End Select
    
	objTCRun.Status=strStatus
    objTCRun.Post
    objTCRun.CopyDesignSteps
    objTCRun.Post	
	strAttachmentFilePath = gobjReportSettings.ReportPath & "\Excel Results\" &gobjReportSettings.ReportName & ".xls"
    Call QC_Upload_Results(objTCRun,strAttachmentFilePath)  
	strAttachmentFilePath = gobjReportSettings.ReportPath & "\HTML Results\" &gobjReportSettings.ReportName & ".html"
    Call QC_Upload_Results(objTCRun,strAttachmentFilePath)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objStartFolder = gobjReportSettings.ReportPath & "\Screenshots"	
	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files	
	For Each objFile in colFiles	
	    strAttachmentFilePath = objStartFolder&"\"&objFile.Name
	    Call QC_Upload_Results(objTCRun,strAttachmentFilePath) 
	Next  
	Set objFolder = Nothing
	Set colFiles =	Nothing 	
End Function

Function UploadTestResultsE2E(objTCRun)
	Select Case g_Flag
    Case 0
                strStatus="Passed"
    Case Else
                strStatus="Failed"
    End Select
    gobjDBCon.Update_Result "Passed", gobjTestParameters.CurrentTestcase 
	objTCRun.Status=strStatus
    objTCRun.Post
    objTCRun.CopyDesignSteps
    objTCRun.Post	
	strAttachmentFilePath = gobjReportSettings.ReportPath & "\Excel Results\" &gobjReportSettings.ReportName & ".xls"
    Call QC_Upload_Results(objTCRun,strAttachmentFilePath)  
	strAttachmentFilePath = gobjReportSettings.ReportPath & "\HTML Results\" &gobjReportSettings.ReportName & ".html"
    Call QC_Upload_Results(objTCRun,strAttachmentFilePath)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objStartFolder = gobjReportSettings.ReportPath & "\Screenshots"	
	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files	
	For Each objFile in colFiles	
	    strAttachmentFilePath = objStartFolder&"\"&objFile.Name
	    Call QC_Upload_Results(objTCRun,strAttachmentFilePath) 
	Next  
	Set objFolder = Nothing
	Set colFiles =	Nothing 	
End Function
	
Function QC_Add_TestRuns()'strAttachmentFilePath)
	Dim intTestCaseID, intTestSetID
    intTestCaseID = Environment("TestCaseID")'Environment.Value("TestID")
    intTestSetID =  Environment("TestSetID")'Environment.Value("TestSetId")

    'strAttachmentFilePath = "H:\run_results.html"
    blnAddToTSet = True
    Select Case g_Flag
    Case 0
                strStatus="Passed"
    Case Else
                strStatus="Failed"
    End Select
    If Isnumeric(intTestCaseID)=False and Isnumeric(intTestSetID) = False Then 'Checks whether Test case Id and Test Set ID is provided in the datafile
                    Reporter.ReportEvent micfail, "Uploading test run to QC failed. Either Test case Id or Test Set ID is not available in the datafile"
                    'Craft Results
                    Exit Function
    End if
   'Set up required QC connections
    Set objQCCon = QCUtil.QCConnection
    Set objTestFact = objQCCon.TestFactory
    Set objTestSetFact = objQCCon.TestSetFactory
    Set objTestitem = objTestFact.item(intTestCaseID)
    Set objTestSetitem = objTestSetFact.item(intTestSetID)
    Set objTestcase = objTestSetitem.TSTestFactory 
    Set objTestPlanFilter=objTestFact.Filter
    objTestPlanFilter("TS_TEST_ID")=intTestCaseID
    Set objTestList=objTestFact.NewList(objTestPlanFilter.Text)
    If objTestList.Count=0 Then
                Reporter.ReportEvent micfail, "Uploading test run to QC failed. Given test case ID not available in test plan"
                'Craft Results
                Exit Function
    End If
    Set objTstSetLst = objTestcase.NewList("") 'Checks whether the given test case is available in the test set
    For Each objTCItem In objTstSetLst
                If Int(objTCItem.field("TC_TEST_ID")) = Int(intTestCaseID)Then
                            blnAddToTSet = "False"
                            Set objTC=objTestcase.newlist(objTCItem.id) 
                            For Each objActTC in objTC
                                        Set objTCtoTS = objActTC
                            Next
                            Exit for
                End If
    Next
    If blnAddToTSet = True Then
                Set objTCtoTS=objTestcase.additem(null)
                objTCtoTS.Field("TC_TEST_ID")=objTestitem
                objTCtoTS.post
    End If
    Set objRunFact=objTCtoTS.RunFactory
    Set objTCRun=objRunFact.additem(objRunFact.UniqueRunName)
    objTCRun.Field("RN_DRAFT") = Environment("Dry_Run")
    objTCRun.Field("RN_RUN_NAME") = gobjTestParameters.CurrentTestcase
    objTCRun.Status=strStatus
    objTCRun.Post
    objTCRun.CopyDesignSteps
    objTCRun.Post
    'Attach execution reports and screenshots to QC run
    strAttachmentFilePath = gobjReportSettings.ReportPath & "\Excel Results\" &gobjReportSettings.ReportName & ".xls"
    Call QC_Upload_Results(objTCRun,strAttachmentFilePath)  
	strAttachmentFilePath = gobjReportSettings.ReportPath & "\HTML Results\" &gobjReportSettings.ReportName & ".html"
    Call QC_Upload_Results(objTCRun,strAttachmentFilePath)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objStartFolder = gobjReportSettings.ReportPath & "\Screenshots"	
	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files	
	For Each objFile in colFiles	
	    strAttachmentFilePath = objStartFolder&"\"&objFile.Name
	    Call QC_Upload_Results(objTCRun,strAttachmentFilePath) 
	Next  
	Set objFolder = Nothing
	Set colFiles =	Nothing 
End Function

'Upload execution results for manual runs
Public Function QC_Upload_Results(objTCRun,strAttachmentFilePath)
'	If (gobjReportSettings.ExcelReport) Then
'            gobjReportSettings.ReportPath & "\Excel Results\" &_gobjReportSettings.ReportName & ".xls"
'     End If
'        
'    If (gobjReportSettings.HtmlReport) Then
'        gobjQCUtility.AttachFileToTestRun gobjReportSettings.ReportPath & "\HTML Results\" &_
'                                                                        gobjReportSettings.ReportName & ".html"
'    End If
'        gobjQCUtility.AttachFolderToTestRun gobjReportSettings.ReportPath & "\Screenshots"
'        
    Set objRnAtt = objTCRun.Attachments.additem(null)
    objRnAtt.FileName= strAttachmentFilePath
    objRnAtt.type=1
    objRnAtt.Description = "Execution Report and screenshots"
    objRnAtt.Post
    objRnAtt.Refresh
    objTCRun.Post  
End Function
	
End Class
'#######################################################################################################################
