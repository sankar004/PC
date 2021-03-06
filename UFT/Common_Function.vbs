

Function setObj(Obj_Browser,Obj_Page,ByRef strObj)
	If Obj_Browser="NA" And Obj_Page ="NA" Then
		If Instr(strObj,"MFO")<>0 Then
			strObj = Replace(strObj,"MFO_","")
	        strQuery = "Select * from [CLA_Repository$] where Obj_Name = '" & strObj & "'"
			Set RSObjID = ExtractMFOID (strQuery)
			strID = RSObjID.Fields(1)
			If IsNumeric(strID) Then
				Set SetObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:="&strID)
			Else
				Set SetObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("attached text:="&strID)
			End If
			Set RSObjID = Nothing
		Else
			setObj = strObj
		End If
		Exit Function 
	Else
	    strType = UCASE(Split(strObj,"_")(0))
	    Set PrntObj = Browser(Obj_Browser).Page(Obj_Page)
	    Select Case strType
	        Case "EDI"
	            set CrntObj = PrntObj.WebEdit(strObj)
	        Case "LNK"
	            set CrntObj = PrntObj.Link(strObj)
	            strObj = "CLK_"&strObj
	        Case "ELE"
	            set CrntObj = PrntObj.WebElement(strObj)
	            strObj = "CLK_"&strObj
	         Case "MH"
	            set CrntObj = PrntObj.WebElement(strObj)
	        Case "BTN"  
	            set CrntObj = PrntObj.WebButton(strObj)
'	            wait 1
	            strObj = "CLK_"&strObj
	        Case "RDO"
	            set CrntObj = PrntObj.WebRadioGroup(strObj)
'	            wait(4)
	        Case "CHK"
	            set CrntObj = PrntObj.WebCheckbox(strObj)
	        Case "LST"
	        
	            set CrntObj = PrntObj.WebList(strObj)
	        Case "MFO"
	        	strObj = Replace("MFO_","")
	        	strQuery = "Select * from [CLA_Repository$] where Obj_Name = '" & strObj & "'"
				Set RSObjID = ExtractSceKeywords (strQuery)
				strID = RSObjID.Fields(1)
'				If IsNumeric(strID) Then
					Set SetObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:="&strID)					
'				Else
'					Set SetObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("attached text:="&strID)
'				End If
				Set RSObjID = Nothing
	        Case Else
	        	set CrntObj = PrntObj
    End Select

' 	strObj_exist = Wait_Obj_Existence(CrntObj)
'   If strObj_exist ="True" Then    	
    	Set setObj = CrntObj
'    End If
  End If 
End Function


Function Wait_Obj_Existence(Act_Obj,strCurrentKeyword)
	incwait = 0
	Flag="false"
	Do
          If Act_Obj.Exist(1) Then
            'If Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).WebElement("ELE_PleaseWait").Exist(1) <> True Then
		        Flag="True"
		        gobjReport.UpdateTestLog strCurrentKeyword & " Screen", "successfully navigated to the page: " &strCurrentKeyword , "PASS"
		        Exit Function
           ' End If
         Else
            Wait 1
		    incwait = incwait+1
            'ObjName = Crntobj.GetROProperty("name")
			'MsgBox "The Object" &ObjName& "does not Exist in current Screen." & vbCrLf & "Make sure you are navigating on correct screen", ,"Object Does Not EXIST"
	    End If
	Loop While (incwait < 20)
	MsgBox "The Object does not Exist in current Screen." & vbCrLf & "Make sure you are navigating on correct screen", ,"Object Does Not EXIST"
	Wait_Obj_Existence=Flag	
End Function



Function Execute_Flow(Obj_Browser, Obj_Page,Execution_Flow)
'	On Error Resume Next
	arrObj = Split(Execution_Flow,";")
	For inc = 0 to Ubound(arrObj)
		strData = GetData(arrObj(inc))
		'Nbv_user=gobjDataTable.GetData("Test_Data", "EDI_USER")
		'msgbox Nbv_user
		If strData <> "Skip"  Then
  		   If Instr(1,UCASE(strData),"UNIQUE") <> 0 Then
			Call UniqueData(strData)
			End If
			If Obj_Browser="NA" And Obj_Page="NA" Then
				if instr (1,Execution_Flow,"CAL")>0 then
					FnName = Replace(arrObj(inc),"CAL_", "")			
					Execute FnName
					Exit function
				End if
			End If		
			Set Act_Obj = setObj(Obj_Browser,Obj_Page,arrObj(inc))
			ObjType = UCASE(Left(arrObj(inc),3))
			If instr (1,Execution_Flow,"Verify")>0 then
				Set Prntobj = Browser(Obj_Browser).Page(Obj_Page)					
				Execute DataVerify(Act_Obj,ObjType,Prntobj,strData)
				Exit function
			
			End If	
			
			If inc = 0 Then
				strObj_exist = Wait_Obj_Existence(Act_Obj)
			End If		
				
			Select Case ObjType
				Case "EDI" ' Webedit
					If Instr(arrObj(inc),"DEV")>0 Then
	                    Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
	                    Act_Obj.Click
	                    myDeviceReplay.SendString strData(inc)
	                Else
	                    Act_Obj.Set strData
	                End if 
				Case "CLK" ' Webbutton
					If instr(arrObj(inc),"Search") Then	
						call MouseOver(Act_Obj)	 
						wait 5
					Else	
						wait 1					
						Act_Obj.Click
					End if	
				Case "LNK"	
						Act_Obj.Click		
				Case "LST"
					
	                If strdata = "Random" Then
	                	Act_Obj.Select(1)
					Else		
						Act_Obj.Select strdata
					End if
					If instr(arrObj(inc),"LST_CPID") Then	
						Act_Obj.Select "#0"	
					End If
				
				Case "CHK"
					If Act_Obj.Exist(1) then
						Act_Obj.Set "ON"
					End If
				Case "RDO"
					Act_Obj.select strdata	
				Case "CAL"
					FnName = Replace(arrObj(inc),ObjType & "_", "")			
					Execute FnName			
			End Select
		End If
	Next
End Function

Function UniqueData(strData)
	strData = Replace(strData,"_Unique","")
	DateTimeStamp = Replace(Now()," ","")
	DateTimeStamp = Replace(DateTimeStamp,"/","")
	DateTimeStamp = Replace(DateTimeStamp,":","")
	DateTimeStamp = Right(DateTimeStamp,4) & Left(DateTimeStamp,Len(DateTimeStamp)-4)
	strData =  DateTimeStamp & strData
	UniqueData = strData
	'call OutputAccountName("CustomerName",strData)
End Function


Function GetData(arrObj)
	ObjType = UCASE(Left(arrObj,3))
	'SheetName = gobjDataTable.GetData("Scenario", "Page")
		If ObjType = "EDI" or ObjType = "LST" or ObjType = "RDO" or ObjType = "CHK"   Then
		   strData = Trim(gobjDataTable.GetData("Test_Data", arrObj))
		   If strData = "" Then
		   		GetData = "Skip"
		   	Else
		   		GetData = strData
		   End If
		   'GetData = gobjDataTable.GetData(SheetName, arrObj)
		  'ElseIf ObjType = "COL" Then
'				Call Fill_UnderWriterQues(arrObj)
		End If

End Function

Function Fill_UnderWriterQues(arrObj)
	RdoFlag = 0
	EdiFlag = 0
	ChkFlag = 0
	arrOpt = Split(arrObj,",")
    Set ODesc = Description.Create
	ODesc("micclass").value = "WebRadioGroup"
	Set arrchildobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
    For inc = 0 To Ubound(arrOpt)
	 If arrOpt(inc) = "NO" or  arrOpt(inc) = "YES" Then
		  
		  arrchildobj(RdoFlag).select arrOpt(inc)
		  RdoFlag = RdoFlag+1
   	  ElseIf arrOpt(inc) = "Test" Then
   	         
		   	  ODesc("micclass").value = "WebEdit"
			  ODesc("visible").value = True
			  Set arrchildobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
			  arrchildobj(EdiFlag).set arrOpt(inc)
			  EdiFlag = EdiFlag +1
	  ElseIf arrOpt(inc) = "ON" Then
   	         
		   	  ODesc("micclass").value = "WebCheckBox"
			  ODesc("visible").value = True
			  Set arrchildobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
			  'For ChkFlag = 0 To arrchildobj.count-1 
			  	arrchildobj(arrchildobj.count-1).set arrOpt(inc)
			  '	ChkFlag = ChkFlag +1
			  'Next
			  
     End If
	Next
End Function
	

Function MouseOver(objmouseover)
    gobjReport.UpdateTestLog "Mouse Over", "Performing mouseover operation on the main menu", "Done"
   'With Browser("Browser_iConnect").Page("Home")
   x=objmouseover.GetROProperty("abs_x")
   y=objmouseover.GetROProperty("abs_y")
   Set obj=CreateObject("Mercury.DeviceReplay")
   obj.MouseMove x+10,y+10
   wait 2
  'obj.MouseClick x+10,y+10,0
'   objmouseover.click
   obj.MouseClick x+10,Y+10,Left_Mouse_Button
   'End With
End Function


Function Fill_Data(Crr_Browser,Crr_Page,Execution_Flow)
    arrObj = Split(Execution_Flow,";")
    Set Prntobj = Browser(Crr_Browser).Page(Crr_Page)
    For inc = 0 to Ubound(arrObj)
        ObjType = Left(arrObj(inc),3)
        If ObjType = "EDI" or ObjType = "LST" or  ObjType = "RDO"Then
                strData = gobjDataTable.GetData("Test_Data", arrObj(inc))
                'arrObj(inc)=Right(arrObj(inc),Len(arrObj(inc))-4)
        End If
        Select Case ObjType
            Case "EDI" ' Webedit
            	Prntobj.WebEdit(arrObj(inc)).Set strData 
				If Instr(arrObj(inc),"EDI_ClassCode") Then
                    Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
                    Prntobj.WebEdit(arrObj(inc)).Click
                    myDeviceReplay.SendString strdata
                    msgbox "Class Code"
                End If
                gobjReport.UpdateTestLog "Text Entered", "Successfully entered the Text in the Field"&arrObj(inc),"Pass"
                'wait 1
            Case "BTN" ' Webbutton
            	If Instr(arrObj(inc),"BTN_Login") Then
            		'msgbox "Login"
            	gobjReport.UpdateTestLog "Login Success", "Successfully Logged into the Application","Pass"
            	End If
'            	If Instr(arrObj(inc),"BTN_ContiM</") Then
'            		msgbox "Continue"
'            	'gobjReport.UpdateTestLog "Login Success", "Successfully Logged into the Application","Pass"
'            	End If
        	   Prntobj.WebButton(arrObj(inc)).Click
                wait 2   
			   gobjReport.UpdateTestLog "Button Clicked", "Successfully Clicked "&arrObj(inc),"Pass"                
            Case "LNK" 'Links
            	If Instr(arrObj(inc),"LNK_LogOff") Then
            		msgbox "Qutoe_ID"
'	            	Quote_ID = Prntobj.WebElement("QuoteID").GetROProperty("innertext")
'	                gobjReport.UpdateTestLog "Quote ID: ", "Quote ID is "&Quote_ID, "Pass"
	            End If
           	   Prntobj.Link(arrObj(inc)).Click
           	   gobjReport.UpdateTestLog "Link Clicked", "Successfully Clicked "&arrObj(inc),"Pass" 
            Case "ELE"
               Prntobj.WebElement(arrObj(inc)).Click
               gobjReport.UpdateTestLog "WebElement Clicked", "Successfully Clicked "&arrObj(inc),"Pass" 
'               If Instr(arrObj(inc),"ELE_QuoteID") Then
'                    Quote_ID = Prntobj.WebList(arrObj(inc)).GetROProperty("innertext")
'                    gobjReport.UpdateTestLog "Quote ID: ", "Quote ID is "&Quote_ID, "Pass"
'                End If
            Case "IMG"
               Prntobj.Image(arrObj(inc)).Click
               gobjReport.UpdateTestLog "Image Clicked", "Successfully Clicked the "&arrObj(inc),"Pass" 
            Case "LST"
'            		If Instr(arrObj(inc),"LST_ProducerCode") Then
'            		msgbox "PC"
'            		End If
                 If Instr(arrObj(inc),"LST_CPID") Then
            		msgbox "CPID"
                    Prntobj.WebList(arrObj(inc)).Select "#0"
                Else If Instr(arrObj(inc),"LST_PolCovLiabLmt") Then
                    Prntobj.WebList(arrObj(inc)).Select "$1,000,000"
                Else
                    Prntobj.WebList(arrObj(inc)).Select strdata
                End If
                End If
              gobjReport.UpdateTestLog "List Value Selected", "Successfully selected a value in the List "&arrObj(inc),"Pass" 
            Case "CHK"
            	
                Prntobj.WebCheckBox(arrObj(inc)).Set "ON"
              gobjReport.UpdateTestLog "Checkbox Checked", "Successfully Checked "&arrObj(inc),"Pass" 
            Case "RDO"
            	Prntobj.WebRadioGroup(arrObj(inc)).select strdata
            gobjReport.UpdateTestLog "Radio Button Selected", "Successfully selected"&arrObj(inc),"Pass" 
            Case "CAL"
            	Call VerifyObj(Prntobj)
            	
            	 
        End Select
'        Wait 2
    Next
    
End Function


Function VerifyObj(Prntobj)
	Quote_ID = Prntobj.WebElement("CAL_QuoteID").GetROProperty("innertext")
    gobjReport.UpdateTestLog "Quote ID: ", "Quote ID is "&Quote_ID, "Pass"
End Function


Function  Verify_ObjExist(strKeyword,arrObj) 		
		Dim incwait, Flag
		incwait = 0
		Set Odesc = Description.Create
			Flag = "Fail"
			If Instr(1,strKeyword,"Navigate")  Then ' To check if the object need to be clicked
				SelectObj = "Click"
'				arrObj(inc) = Replace(arrObj,"_Click","")
			ElseIf Instr(1,arrObj,"_Neg")  Then
				SelectObj = "Neg"
				arrObj = Replace(arrObj,"_Neg","")
			End If
			
			Set obj = SetObj(arrObj)
				If obj.Exist(1) Then
					gobjDataTable.PutData "HIMCO_Data","Status","Pass"
					'obj.highlight
					Flag = "Pass"
					If SelectObj  = "Neg" Then
						
						gobjReport.UpdateTestLog "Verification for errors", "Errors are observed in the current page", "Fail"
						gobjDataTable.PutData "HIMCO_Data","Status","Fail"
					End If
					
					If SelectObj = "Click" Then
						
						strlink_data = gobjDataTable.GetData("HIMCO_Data","Screen_Navigate")
						arrlink_obj = Split(strlink_data,",")
						objlink_cnt = UBound(arrlink_obj)		
							If arrObj<>arrlink_obj(objlink_cnt) Then
								Call MouseOver(obj,arrObj)
							Else
'								obj.highlight
								ClickTime1=time()
								obj.Click
								x=obj.GetROProperty("abs_x")
							    y=obj.GetROProperty("abs_y")
							    Set Devobj=CreateObject("Mercury.DeviceReplay")
							    Devobj.MouseMove x+20,y+20
								
							End If
					
						Set Popup = Browser("micclass:=Browser").Dialog("micclass:=Dialog")
							If Popup.Exist(1) Then
								Call PopupCheck(arrObj,Popup)
								Exit Function
							End If
'					
						gobjReport.UpdateTestLog "Click "&arrObj, arrObj& " is selected", "Done"
					Else
						ClickTime2=time()
						Call TimeDifference(ClickTime2,ClickTime1,arrObj)
						gobjReport.UpdateTestLog "Verify "&arrObj, arrObj& " is available in the application. The Page loaded successfully", "Pass"
					
					End If
			Else 
				Wait 1
				incwait = incwait + 1
			End If		
					If Flag = "Fail" and SelectObj   = "Neg" Then
						gobjReport.UpdateTestLog "Verification for errors", "No Error is observed in the current page", "Pass"			
					ElseIf Flag = "Fail" Then
						gobjReport.UpdateTestLog "Verification for Link ", "The Link is not available in the application or The page is taking more than 5 Seconds to load", "Fail"	
						gobjDataTable.PutData "HIMCO_Data","Status","Fail"
						
					End If
End Function


Function ExecuteAction_WebPage(RstKeywords)
	Crr_Browser = RstKeywords.Fields(1)
	Crr_Page = RstKeywords.Fields(2)
	Strfields = RstKeywords.Fields.count
	For incfield = 3 to Strfields-1
		Execution_Flow = RstKeywords.Fields(incfield)
		If Execution_Flow <> "" Then
			Call Execute_Flow(Crr_Browser,Crr_Page,Execution_Flow)
			gobjReport.UpdateTestLog "Success: Passed", "Test Iteration Executed successfully", "Passed"
		Else
			Exit For
		End If
		
	Next
'	wait 2
End Function

Function ExecuteAction_Mainframe(RstKeywords)
	Strfields = RstKeywords.Fields.count
	For incfield = 3 to Strfields-1
		Execution_Flow = RstKeywords.Fields(incfield)
		If Execution_Flow <> "" Then
			arrObj = Split(Execution_Flow,",")	
			For inc = 0 to Ubound(arrObj)			
				If arrObj(inc)= "NxtScreen" Then
					TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
					TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
			    ElseIf arrObj(inc)= "Final" Then
		        	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_PF9  
		        ElseIf arrObj(inc)= "RISK_CA" Then
		        	Call RISK_CA()  
				Else
					arrObj(inc) = "MFO_" & arrObj(inc)
					Set TstObj = Setobj("NA","NA",arrObj(inc))
				    ' TstObj.highlight
				 	    
				     strdata = gobjDataTable.GetData("CLA_Data",arrObj(inc))	
				    If arrObj(inc) = "CPID_WC" Then		  
							    
				    	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
						TstObj.Set strdata
					Else
						TstObj.Set strdata
				    End If		 
				End If  
			Next
		End If
		
	Next
End Function
	
Function ExecuteFlow(RstKeywords)
	strAppType = Left(RstKeywords.fields(0),3)
		Select Case strAppType 
			Case "CLA"
				Call ExecuteAction_Mainframe(RstKeywords)
			Case Else
				Call ExecuteAction_WebPage(RstKeywords)
		End Select
End Function



Function GetBrowsers()
	Set Odesc = Description.Create	
	Odesc("micclass").Value = "Browser"
	Set arrBrowsers = Desktop.ChildObjects(Odesc)
	Set GetBrowsers = arrBrowsers
End Function






Function ExtractMFOID (strQuery)
		'CheckPreRequisites()
		Set adodb = CreateObject("ADODB.Connection")
		adodb.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
		"Data Source=C:\Users\COG11704\Desktop\IAS1\Datatables\CLA_Compass_Automation_Suite.xls;" & _
		"Extended Properties=""Excel 8.0;HDR=Yes"";" 
		Set objSceKeywords = CreateObject("ADODB.Recordset")
		Set objSceKeywords = adodb.Execute(strQuery)
		Set ExtractMFOID= objSceKeywords
			Set adodb = nothing
		Set objSceKeywords = nothing
	End Function
	
	Function DataVerify(Act_Obj,ObjType,Prntobj,strData)
		Select Case ObjType
			Case "Edi" ' Webedit
				VerifyData = Act_Obj.GetROProperty("value")
				If Instr(1,VerifyData,strData)<>0 Then
					gobjReport.UpdateTestLog "Verify Edit Box", "The value "&VerifyData&" is displayed as expected", "Pass"
				Else
					gobjReport.UpdateTestLog "Verify Edit Box", "The value "&VerifyData&" is not displayed as expected", "Fail"
				End If		
				
			Case "Lnk" 
				Set Odesc = Description.Create
				Odesc("micclass").Value = "Link"
				Odesc("name").Value = arrObj(inc)
				Odesc("visible").Value = "True"
				Set arrLink = Prntobj.ChildObjects("Odesc")
				If arrLink.count>0 Then
					VerifyLink = arrLink(0).GetROProperty("name")
					If Instr(1,VerifyLink,strData)<>0 Then
						gobjReport.UpdateTestLog "Verify Link", "The value "&VerifyLink&" is displayed as expected", "Pass"
					Else
						gobjReport.UpdateTestLog "Verify Link", "The value "&VerifyLink&" is not displayed as expected", "Fail"
					End If
				End If
				
			Case "Ele"
				Set Odesc = Description.Create
				Odesc("micclass").Value = "WebElement"
				Odesc("name").Value = arrObj(inc)
				Odesc("visible").Value = "True"
				Set arrEle = Prntobj.ChildObjects("Odesc")
				If arrEle.count>0 Then
					VerifyText = arrEle(0).GetROProperty("innertext")
					If Instr(1,VerifyText,strData)<>0 Then
						gobjReport.UpdateTestLog "Verify Link", "The value "&VerifyText&" is displayed as expected", "Pass"
					Else
						gobjReport.UpdateTestLog "Verify Link", "The value "&VerifyText&" is not displayed as expected", "Fail"
					End If
				End If
			
			Case "Lst" 
				VerifyLstData = Act_Obj.GetROProperty("value")
				If Instr(1,VerifyLstData,strData)<>0 Then
					gobjReport.UpdateTestLog "Verify List Value", "The value "&VerifyLstData&" is displayed as expected", "Pass"
				Else
					gobjReport.UpdateTestLog "Verify List Value", "The value "&VerifyLstData&" is not displayed as expected", "Fail"
				End If
					

		End Select
		Wait 3
End Function

Function CaptureSnapshot()
	'path = "C:\Users\COG12193\Desktop\Compas Automation\Test Data for PC\SnapShots\"
	ScreenshotPath = "H:\Screenshot\"& gobjTestParameters.CurrentTestcase
	dim filesys, newfolder, newfolderpath 
    'ScreenshotPath = path&FolderName
    set filesys=CreateObject("Scripting.FileSystemObject") 
    If Not filesys.FolderExists(ScreenshotPath) Then 
    Set newfolder = filesys.CreateFolder(ScreenshotPath) 
'    msgbox "Folder has been created"
    Else
'    msgbox "Folder is available"
    End If
    
    Dim ScreenName
      On Error Resume Next
'    ScreenName = "GenInfo"
    CurrentTime = "Screenshot"&"_"& Day(Now)&"_"& Month(Now)&"_"& Year(Now)&"_"& Hour(Now)&"_"& Minute(Now)&"_"& Second(Now)
    'Set the screen shot name
    ScreenShotName = CurrentTime & ".png"
    'Final screenshot location
    ScreenName = ScreenshotPath&"\"&ScreenShotName
    ' just capture
    Desktop.CaptureBitmap ScreenName,True
 End Function
 
' Function LoginECOS_1
'	Dim oDesc, x
'	strflag = "Fail"
'	Set oDesc = Description.Create
'	oDesc( "micclass" ).Value = "Browser"	 
'	If Desktop.ChildObjects(oDesc).Count > 0 Then
'	    For x = Desktop.ChildObjects(oDesc).Count - 1 To 0 Step -1
'	       If InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"Quality Center") = 0 Then  
'	          Browser( "creationtime:=" & x ).Close
'	        ElseIf InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"ECOS")  <> 0 Then 
'				strflag = "Pass"	        
'	       End If
'	    Next
'	End If
'	If strflag <> "Pass" Then  
'     strURL=gobjDataTable.GetData("General_Data","Url")
'     
'     Systemutil.Run "iexplore",strURL
'    End If
'    
'    call Userlogin_1(strURL)	
' End Function

Sub Userlogin_1(strURL)
    wait 2
	With Browser("ECOS").Page("Login")
		If .WebEdit("EDI_USER").Exist(1) Then
			Ecos_user=gobjDataTable.GetData("General_Data","EDI_USER")
			Ecos_pwd=gobjDataTable.GetData("General_Data","EDI_PASSWORD")
            .WebEdit("EDI_USER").Set Ecos_user
			.WebEdit("EDI_PASSWORD").Set Ecos_pwd
			.WebElement("ELE_Login").Click
			Exit Sub
		
		End If

	End With	
End Sub

'###################################################################################################################
'Function Description   : Function to Fill the expected PolicyID in Data sheet 
'Input Parameters		: strFieldName,strDataValue
'Author					: Shankar
'Date Created			: 23/09/2015
'###################################################################################################################
Function Outputdata(strFieldName,strDataValue)


		Dim strFilePath, strConnectionString, r_objConn,Cmd, strQuery
'        strFilePath = gobjExcelDataAccess.m_strDatabasePath & "\Common Testdata.xls"
		strFilePath =	"C:\E2E\E2E\ETE_Scenarios.xls"
        m_intCurrentIteration=Environment("TestIteration")
        m_intCurrentSubIteration=1
        m_strCurrentTestCase=Environment("TestName")
        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
		
		Dim strNonQuery, objTestData
		strTestDataSheet = "EndToEnd"


		strQuery = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = '" & strDataValue & "'" &_
																" where TC_ID = '" & m_strCurrentTestCase &_
																"' and Iteration = " & m_intCurrentIteration & ""
																
															
				r_objConn.Execute strQuery                       
        '
        r_objConn.Close
		'Set objTestData.ActiveConnection = nothing
		gobjReport.UpdateTestLog "Output value", _
								"Output value " & strDataValue & " written into the " & strFieldName & " column", "Done"
End Function

Function ExecuteQuery(strQuery)
Dim strFilePath, strConnectionString, r_objConn,Cmd
        strFilePath = Environment("OutputFileName")
            strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConnect = CreateObject("ADODB.Connection")
        r_objConnect.Open strConnectionString
		
		Dim strNonQuery, objTestData
		strTestDataSheet = "EndToEnd"


		strQuerytest=strQuery
															
				Set objTestData = CreateObject("ADODB.Recordset")													
		 Set Cmd = CreateObject("ADODB.Command")
        Cmd.ActiveConnection = r_objConnect
        Cmd.CommandText = strQuerytest
        'Dim objTestData: Set objTestData = CreateObject("ADODB.Recordset")
        objTestData.CursorLocation = 3
      objTestData.Open Cmd  
    If instr(strQuerytest,"Update")>0Then
    QueryResult="Completed"
      	
      Else
      QueryResult  = objTestData.Fields(0)
    End If
      
       r_objConnect.Close   
        ExecuteQuery=QueryResult
		'Set objTestData.ActiveConnection = nothing
End Function
Function Login_CLA()
	'closeBrowsers()
	If TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Exist Then
	
		Msgbox "CLA is available"
		
	Else
	strURL= "http://hodcorp/hod/g-c-hocl.html?JavaType=java2"
	SystemUtil.Run "iexplore.exe" , strURL ,,,3
	Wait 5
	'Browser("Browser").Page("g-c-hocl").Frame("g-c-hocl").ActiveX("ActiveX").WinObject("SunAwtCanvas").DblClick 76,101
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:=1848").Set "IMSW"
	
	
'	TeWindow("TeWindow").TeScreen("screen23111").TeField("APPL").Set "IMSW"
'	TeWindow("TeWindow").TeScreen("screen23111").TeField("APPL").SetCursorPos 4
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
	'Wait 2
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:=361").Set "COG18094"
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:=521").SetSecure "58ff43f46331ebc1e2c8e5e0237febb4358e9096658c"
	'TeWindow("TeWindow").TeScreen("DFS2002 08:04:30 TERMINAL").TeField("NEWPW").SetCursorPos
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER	
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_CLEAR
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey "/Exit"
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey "/Test Mfs"
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey "/For Clalong"
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
	'TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_TAB
	'TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey Julian
	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER



'	Wait 2
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:=78").SetCursorPos 4
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_CLEAR
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey "/test mfs"
'
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SetCursorPos 1,10
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
'	Wait 2
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
'	TeWindow("TeWindow").TeScreen("DFS058I 08:04:54 TEST").TeField("field id:=162").SetCursorPos
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_CLEAR
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey "/for clalong"'clalong"
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SetCursorPos 1,13
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
'	Wait 2
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:=2").SetCursorPos
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
'	TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").Sync		
	End If	
End Function

Function GetDataDetails(Obj,strData)
	'strWriteData = strData
'''''''	 If Instr(1,strData,",") Then
'''''''		arrData = Split(strData,",")
'''''''	End If
'''''''	
'''''''	If Ubound(arrData) <> 0 Then
'''''''		strData = arrData(0)
'''''''	End If
'''''''
'''''''			If strData = "Random" Then
'''''''				strData = Random_Data(Left(Obj,3))
'''''''		   End If
'		   If Instr(1,strData,"*") <> 0 Then
'				strData = Replace(strData,"*","")
'				Call CommonData(Obj,strData)
'			End If
'				strData = strWriteData
'				If Obj = strWriteData Then
'					strWriteData = OutputData(strWriteData)
'					strData = "Skip"
'				End If
'				Call OutputCommondata(Obj,strWriteData)
'				End If 
'			If Instr(1,strData,"*") <> 0 Then
'				strWriteData = Replace(strData,"*","")
'				strData = strWriteData
'				If Obj = strWriteData Then
'					strWriteData = OutputData(strWriteData)
'					strData = "Skip"
'				End If
'				Call OutputCommondata(Obj,strWriteData)
'				End If 
'				arrObj = "Skip"
'			ElseIf Instr(1,Obj,"$") <> 0 Then
'				'Call VerifyData(arrObj,Replace(strData,"$",""))
'				Call VerifyData(Replace(Obj,"$",""))  
'				arrObj = "Skip"
'''''			If Instr(1,strData,"#") <> 0 Then
'''''				strData =  Getcommonoutdata(Replace(strData,"#",""))
'''''			End If
'
			GetDataDetails = strData
End Function

Function OutputCommondata(strFieldName,strDataValue)
	
	Dim strFilePath, strConnectionString, r_objConn,Cmd, strQuery
		strFilePath =	Environment("OutputFileName")
	        m_intCurrentIteration=Environment("TestIteration")
	        m_intCurrentSubIteration=1
	        m_strCurrentTestCase=gobjTestParameters.CurrentTestcase
        

        strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=2"""    
        Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
		
		Dim strNonQuery, objTestData
		strTestDataSheet = "EndToEnd"
		
        strQuery = "Select E2E_Fields from [Field_Mapping$] where NBV = '" &strFieldName& "'"
        
        strFieldName = ExecuteQuery(strQuery) 

		strQuery = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = '" & strDataValue & "'" &_
																" where TC_ID = '" & m_strCurrentTestCase &_
                                                                "' and Iteration = " & m_intCurrentIteration & "" 
															
		r_objConn.Execute strQuery 
		
        r_objConn.Close
		gobjReport.UpdateTestLog "Output value", _
								"Output value " & strDataValue & " written into the " & strFieldName & " column", "Done"
	
End Function

Function OutputCommondata_CLA(Filed_Name,Field_Text,strCurrentKeyword)
	
	Dim strFilePath, strConnectionString, r_objConn,Cmd, strQuery
		strFilePath = Environment.value("DataFilePath")
	        m_intCurrentIteration=Environment("TestIteration")
	        'm_intCurrentSubIteration=1
	        m_strCurrentTestCase=gobjTestParameters.CurrentTestcase
            Subiteration = Environment("Subiteration_Current")
           strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
		
		Dim strNonQuery, objTestData
		strTestDataSheet = strCurrentKeyword
        strFieldName = Filed_Name
      
        	
         		'strQuery = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = "&Field_Text&""&_
																'" where TC_ID = '" & m_strCurrentTestCase &_
                                                                 ' "' and Iteration = " & m_intCurrentIteration &_
                                                					'" and SubIteration = " & Subiteration & ""
	
				strQuery = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = '" & Field_Text & "'" &_
																" where TC_ID = '" & m_strCurrentTestCase &_
                                                                 "' and Iteration = " & m_intCurrentIteration &_
                                                					" and SubIteration = " & Subiteration & ""
      
		r_objConn.Execute strQuery
		r_objConn.Close
		gobjReport.UpdateTestLog "Output value", _
								"Output value " & Field_Text & " written into the " & strFieldName & " column", "Done"		
		      
	
End Function


'###################################################################################################################
'Function Description   : Function to fill the Random value to validate the field
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 09/11/2015
'###################################################################################################################
Function Random_Data(ObjType)
	Select Case ObjType
	    Case "EDI" ' Webedit
		     Random_Data =  "Test"
	    Case "LST"
		     Random_Data = 1
		Case "CHK"
			 Random_Data = "ON"
		Case "RDO"
				'GetRO of All items
				'Split with demimiter
				'Pick up last value
				'Random_Data = "Last Value"
			    ' Obj.select strData(inc)	
	End Select
	
End Function

'Function LoginECOS
'	Dim oDesc, x
'	strflag = "Fail"
'	Set oDesc = Description.Create
'	oDesc( "micclass" ).Value = "Browser"	 
'	If Desktop.ChildObjects(oDesc).Count > 0 Then
'	    For x = Desktop.ChildObjects(oDesc).Count - 1 To 0 Step -1
'	       If InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"Quality Center") = 0 and InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"ICON | Small Commercial Quoting")  = 0 Then  
'	          Browser( "creationtime:=" & x ).Close
'	        ElseIf InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"ECOS")  <> 0 Then 
'				strflag = "Pass"	        
'	       End If
'	    Next
'	End If
'	If strflag <> "Pass" Then  
'     strURL = "http://sit2-ecos.thehartford.com/cc/ClaimCenter.do"
'End If
'Systemutil.Run "iexplore",strURL
'	
'End Function
'###################################################################################################################
'Function Description   : Function to fill the fields after fetching data from data sheet
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 06/10/2015
'###################################################################################################################

Function VehicleEntry(coloumn)
   Set Connection = createobject("ADODB.Connection")
   Set Command = Createobject("ADODB.Command")
   strFilePath =  Environment.value("DataFilePath")
   Vehicles    = Environment("Vehicle")
   strTestDataSheet = "VEHTYPES_Add"
   Con_String =  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2""" 
   strQuery = "Select " & coloumn & " from [" & strTestDataSheet & "$]" &_
                                                " where ID = " & Vehicles &""                                                         
   Connection.open Con_String                                                                                                                
   Command.ActiveConnection = Con_String
   Command.CommandText = strQuery
   Set RecordSet = Command.Execute()
   If coloumn = "Object" Then
   	   VehicleEntry = RecordSet.Fields("Object").Value
   ElseIf coloumn = "TestData" Then
   		 VehicleEntry = RecordSet.Fields("TestData").Value
   End If

End Function
Function DataGet(strCurrentKeyword,obj)
	 Set Connection = createobject("ADODB.Connection")
	 Set objField = CreateObject("ADODB.Command")
	 strFilePath =  Environment.value("DataFilePath")
	 strTestDataSheet = strCurrentKeyword
	 strFieldName = obj
	 m_intCurrentIteration=Environment("TestIteration")
	 m_strCurrentTestCase=gobjTestParameters.CurrentTestcase
	 Con_String =  "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2""" 
     Connection.Open Con_String
     strQuery = "Select " & strFieldName & " from [" & strTestDataSheet & "$]" &_
                                                " where TC_ID = '" & m_strCurrentTestCase &_
                                                  "' and Iteration = " & m_intCurrentIteration &""
     objField.ActiveConnection = Con_String
     objField.CommandText = strQuery
     Set Record = objField.Execute()
     If obj = "Browser" Then
     	DataGet = Record.Fields("Browser").Value
     Else 
     	DataGet = Record.Fields("Object").Value
     End If
     
    
End Function
'Function ObjectProdRip_Flow(strCurrentKeyword)
'	 Object = DataGet(strCurrentKeyword, "Object")
'	 arrObj = Split(Object,";")
'	 IF arrObj(0) = "MFO__NxtScreen" Then
'		TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
'     Else 
'	 For inc = 0 to Ubound(arrObj)
'	  Set Act_Obj = setObjType(arrObj(inc))
'	  ObjType = UCASE(Left(arrObj(inc),3))
'	   Select Case ObjType
'	    Case "MFO"
'	        Obj = Split(arrObj(inc),"_")(1)
'	   		Field_Text = Act_Obj.GetROProperty("text")
'			Filed_Name = Split(arrObj(inc),"_")(2)
'			If Field_Text <> "" Then 
'				call OutputCommondata_CLA(Filed_Name,Field_Text,strCurrentKeyword)
'			End IF 
'        IF arrObj(inc) = "MFO__NxtScreen" Then
'		    TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
'		 End IF 
'	    End Select 	  
'	  Next 
'	  End IF 
'End Function

Function Object_Flow(strCurrentKeyword)
'On error resume next
Environment("Script")= gobjTestParameters.CurrentScenario
If (Instr(1,UCASE(strCurrentKeyword),"CAL")) = 0 Then
    Environment("Current_Scenario") = Mid(strCurrentKeyword,5,8)
    If  Environment("Current_Scenario") = "VEHTYPES" Then
       Vehciles = Split(strCurrentKeyword,";")
    For inc = 1 To ubound(Vehciles)
        Environment("Vehicle") = Vehciles(inc)
        Object = "Object"
        TestData = "TestData"
    	Object = VehicleEntry(Object)
    	TestData = VehicleEntry(TestData)
    	Call Dataset(Object,TestData)
    Next    	
    Else
	   Object = gobjDataTable.GetData(strCurrentKeyword, "Object")
	   TestData = gobjDataTable.GetData(strCurrentKeyword, "TestData")
	End IF 
	arrObj = Split(Object,";")
	arrCrntTestData = Split(TestData,";")
	For inc = 0 to Ubound(arrObj)
	strData = Split(arrCrntTestData(inc),"$$")
		If Instr(strData(0),"!") Then                        '******'For Validating UI fields
			Call UI_E2Evalidation(strData(0),arrObj(inc))
		Else	
		
	Set Act_Obj = setObjType(arrObj(inc))
		If inc=0  Then
		 call Wait_Obj_Existence(Act_Obj,strCurrentKeyword)
		 End If
			
	ObjType = UCASE(Left(arrObj(inc),3))
		If arrObj(inc) <> "Skip" and strData(0) <> "Skip" Then
			If Instr(1,strData(0),"Unique") <> 0 Then
				'strData(0) = UniqueData(strData(0))
				strData(0) = GenerateRandomString(18,"Text")
			End If	
		If Instr(1,strData(0),"#Edi_PolNumber") <> 0 Then
			strData(0) = Getoutdata(strData(0))
			Browser("ICONSmallCommercial").Page("NewCustomer").WebEdit("EDI_Search").Set strData(0)
			Browser("ICONSmallCommercial").Page("NewCustomer").WebEdit("EDI_Search").Click
			Dim Wshell
			set Wshell = CreateObject("WScript.shell")
			Wshell.SendKeys "{ENTER}"
			Browser("ICONSmallCommercial").Page("NewCustomer").Link("LNK_Search").Click
		End If						
					Select Case ObjType
					
						Case "EDI" ' Webedit
							If Instr(arrObj(inc),"DEV")>0 Then
			                    Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
			                    wait 1
			                    Act_Obj.Click
			                    myDeviceReplay.SendString strData(0)
			                    'myDeviceReplay.SendString strData(0)
			                     Wait 3
			                    ' myDeviceReplay.SendString strData(0)
			                    myDeviceReplay.presskey 15
			                    gobjReport.UpdateTestLog arrObj(inc) & " Field", arrObj(inc) & " Field is updated with the data" & strData(0) , "Done"
			               ' Elseif arrObj(inc)="EDI_NPN"  Then
			                   ' Act_Obj.Set strData
			                 
			                    Else
			                    'strdata1=split(strdata(inc),",")
			                    Act_Obj.Set strdata(0)
'			                    If Ubound(strdata1)<> 0 then
'			                    If strdata1(0) <> strdata1(1) Then
'			                    	Act_Obj.Set strdata1(0)
'			                    End If
'			                    Else 
'			                    	Act_Obj.Set strdata1(0)			                    
'			                    End If
			                     
'			                    If strCurrentKeyword = "NewCustomer" Then
'			                    	Outputdata arrObj(inc),strData(inc)
'			                    End If
			                    
			                End if 
						Case "CLK" ' Webbutton
						    'Call CaptureSnapshot()
							If instr(arrObj(inc),"Search") Then	
								call MouseOver(Act_Obj)	 
								ElseIf instr(arrObj(inc),"ReservePolicyNumber") Then	
						 		Act_Obj.Click
						 		Wait 4
								Environment("RO")=Browser("[TEST mode - 8.0.4.538]").Page("PolicyInfo").WebElement("ELE_RO").GetROProperty("innertext")
								Environment("Pol_Sym")=Browser("[TEST mode - 8.0.4.538]").Page("PolicyInfo").WebEdit("EDI_PolicySymbol").GetROProperty("value")
								Environment("Pol_Num")=Browser("[TEST mode - 8.0.4.538]").Page("PolicyInfo").WebEdit("EDI_PolicyNumber").GetROProperty("value")
						ElseIf instr(arrObj(inc),"ELE_C") then
						 		x=Act_Obj.GetROProperty("abs_x")
							    y=Act_Obj.GetROProperty("abs_y")
							    Set Devobj=CreateObject("Mercury.DeviceReplay")
							    'Devobj.MouseMove x+20,y+20
								  Devobj.MouseClick x+10,Y+10,Left_Mouse_Button
						
	'                            wait 5							
								Else	
								if Act_Obj.Exist(30) then
						     	  	Act_Obj.Click	
						     	Else  
									Act_Obj.Click		
								End If
								gobjReport.UpdateTestLog arrObj(inc) & " Field", arrObj(inc) & " Field is selected", "Done"
							End if	
						Case "LNK"	
						     ' Call CaptureSnapshot()
						     	if Act_Obj.Exist(30) then
						     	  	Act_Obj.Click	
						     	Else  
									Act_Obj.Click		
								End If
						Case "LST"
'				            If strData(inc) = "Random" Then
'			                	Act_Obj.Select(1)
'							Else	
								Print "arrObj(inc): " &arrObj(inc)
                               If instr(1,strData(0),".*") Then
                               	   Call AddLocation(strData(0))
                               	   ElseIf arrObj(inc)="LST_ClassDescription_CC" Then
                               	   Call ValidateListItems(strData(0))
                               ElseIf strdata(0) = "Random" Then
	                	              Act_Obj.Select "#2"
	                	           
'                               arrObj(inc)="LST_EmpLiabLimits" OR arrObj(inc)="LST_Location_CC"  Then
'                               	Act_Obj.Select(strData(0))
'                               	ElseIf arrObj(inc)="Lst_EPLIDed" Then
'                               	      Act_Obj.Select "$5,000"
                               Else
                               	wait 2
                               	'strdata2=split(strdata(inc),",")
                               	Act_Obj.Select(strData(0))
                               	
                               
								'Act_Obj.Select(strData(inc))
								End If
'								If strCurrentKeyword = "NewCustomer" Then
'			                    	Outputdata arrObj(inc),strData(inc)
'			                    End If
'								Call OutputCommondata(arrObj(inc))
'							End if
							gobjReport.UpdateTestLog arrObj(inc) & " Field", arrObj(inc) & " Field is updated with the data" & strData(0) , "Done"
								
						Case "CHK"
							If Act_Obj.Exist(1) then
								Act_Obj.Set "ON"
							End If
						Case "RDO"
							If strData(0) = "Click" Then
								Act_Obj.Click								
							Else 
								Act_Obj.select strData(0)							
							End If								
						Case "ELE"
								Act_Obj.Click
	                 		
	                 	Case "IMG"
	                 			if Act_Obj.Exist(30) then
						     	  	Act_Obj.Click	
						     	Else  
									Act_Obj.Click		
								End If
         				Case "MH"
         							Act_Obj.FireEvent "onmouseover"
						Case "FLE"
						     Act_Obj.Set strData(0)
						Case "SRH"
						     Act_Obj.Click
						Case "CAL"
							FnName = Replace(arrObj(inc),ObjType & "_", "")			
							Execute FnName
							gobjReport.UpdateTestLog "Function Call", "Calling function " & FnName , "Done"
							
						Case "EXP"
							
						Browser("[TEST mode - 8.0.4.538]").Page("[TEST mode - 8.0.4.538]").WebElement("innertext:="&strData(0)).Click	

						Case "MFO"
							
							If strdata(0)= "NxtScreen" Then
								TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
								'TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER
								
							ElseIf strdata(0)= "Nxtlob" Then
								TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_PF8
								
							ElseIf strdata(0)= "Save" Then
								TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_PF4
								
			    			ElseIf strdata(0)= "Final" Then
								TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_PF9
								
							ElseIf strdata(0)= "Review" Then
								TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_PF3
							
							ElseIf strdata(0)= "getpolicynumber" Then
							strPolicynumber = Act_Obj.GetROProperty("text")
							CLAPolicynumberupdate (strPolicynumber)
							
							ElseIf strdata(0)= "getppremium" Then
							strPolicynumber = Act_Obj.GetROProperty("text")
							CLAPolicynumberupdate (strPolicynumber)
							
							ElseIf strdata(0) <> "Click" Then							
								If strdata(0)= "Env_Val" Then
									strdata(0) = Getenvironmentvalue(arrObj(inc))									
								End If
								'gobjDataTable.PutData "Summery", "PolicyNo" , "ABC"
								IF Environment("Script") = "CLA_ComPAS_ProdRipOff" Then 
								'If (Instr(1,Environment("TestName"),"ProdRipoff")) <> 0 Then
								  IF strCurrentKeyword = "CLA_Home_Screen" Then 
								     Act_Obj.Set strdata(0)
								  Else 
								     Field_Text = Act_Obj.GetROProperty("text")
								     Filed_Name = Split(arrObj(inc),"_")(2)
								       If Field_Text <> "" Then 
								         call OutputCommondata_CLA(Filed_Name,Field_Text,strCurrentKeyword)
								        End IF 
							       End IF 
								Else 
								   Act_Obj.Set strdata(0)
								End IF 
							    gobjReport.UpdateTestLog arrObj(inc) & " Field", arrObj(inc) & " Field is updated with the data" & strData(0) , "Done"						     
							End If
					End Select
				End If
			arrVerify = Split(arrCrntTestData(inc),",")
				If Ubound(arrVerify) <> 0 Then
					Call VerifyObject(arrVerify)
				End If
		End If	
		Next
	
		
'		If left(strCurrentKeyword,3) = "CLA" then		
'		TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").SendKey TE_ENTER		
'		End if
		
		
Else
	FnName = Replace(strCurrentKeyword,Left(strCurrentKeyword,4), "")			
	Execute FnName
End If

End Function

Function Getenvironmentvalue(objname)
	
	Policy_Field = Split(objname,"_")							
	Select case (Policy_Field(Ubound(Policy_Field)))									
		Case "OFFICE"
		strdata = Environment("RO")
		Case "SYM"
		strdata = Environment("Pol_Sym")
		Case "ADDR"	
		strdata = Environment("Pol_Num")	
		Case "DATE"	
		strdata = Environment("EffDate")
		Case "RnlDATE"	
		strdata = Environment("ExpDate")
		Case "REWRITERO"	
		strdata = Environment("RWRO")
		Case "REWRITESYM"
		strdata = Environment("RWPol_Sym")
		Case "REWRITENUM"		
		strdata = Environment("RWPol_Num")
		Case "RECEIVED"	
		strdata = Environment("ReceivedDate")
		
	End Select
	Getenvironmentvalue = strdata
End Function

Function CLAPolicynumberupdate (strPolicynumber)
	
sFileName = "C:\Users\co24262\Desktop\PolicyNumber.xls"
'strDataValue = "CLA_UM_AL"
strTCID = gobjTestParameters.CurrentTestcase
strFieldName ="PolicyNumber"
strDataValue = strPolicynumber
strTestDataSheet = "Sheet1"

		Set objConnection = CreateObject("ADODB.Connection")
		objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sFileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;"";"
		strSQL = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = '" & strDataValue & "'" &_
                                                                " where TC_ID = '" & strTCID &"'"
                                                                
		'strSQL = "INSERT INTO [" & strTestDataSheet & "$](TC_ID, PolicyNumber) VALUES ('"&strTCID"','" &strDataValue "')"                                                             
        objConnection.Execute(strSQL)
        objConnection.Close        
        
End Function


Function Launch_AccountRecepient()
	Set policyactions= Browser("name:=.*").Page("title:=.*").WebElement("innertext:=Actions","html id:=.*MenuActions-btnInnerEl")
	policyactions.Highlight
	policyactions.Click
	
	Set objAccountRecepient=Browser("name:=.*").Page("title:=.*").WebElement("html id:=.*AccountFile:AccountFileMenuActions:AccountFileMenuActions_Create:AccountFileMenuActions_NewAccountRecipient-textEl","innertext:=New Account Recipient")
	If Browser("name:=.*").Page("title:=.*").WebElement("html id:=.*AccountFile:AccountFileMenuActions:AccountFileMenuActions_Create:AccountFileMenuActions_NewAccountRecipient-textEl","innertext:=New Account Recipient").Exist Then
	objAccountRecepient.Highlight
	objAccountRecepient.Click
	End If 
End Function
