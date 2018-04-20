	
'###################################################################################################################
'Function Description   : Function check for the Browser name calls a Function userlogin()
'Input Parameters		: None
'Return Value    		: call userlogin()
'Author					: Shankar
'Date Created			: 19/09/2015
'###################################################################################################################
Sub LoginNBV()
	'QCUtil.CurrentRun.Field("RN_DRAFT") = "N"
'	On Error Resume Next
	Dim oDesc, x
	strflag = "Fail"
	'Create a description object
	Set oDesc = Description.Create
	oDesc( "micclass" ).Value = "Browser"	 
	'Close all browsers except Quality Center
	If Desktop.ChildObjects(oDesc).Count > 0 Then
	    For x = Desktop.ChildObjects(oDesc).Count - 1 To 0 Step -1
	       If InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"Quality Center") = 0 and InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"ICON | Small Commercial Quoting")  = 0 Then  
	          Browser( "creationtime:=" & x ).Close
	        ElseIf InStr(1,Browser("creationtime:="&x).GetROProperty("name"),"ICON | Small Commercial Quoting")  <> 0 Then 
				strflag = "Pass"	        
	       End If
	    Next
	End If
	If strflag <> "Pass" Then  
	  strURL=gobjDataTable.GetData("General_Data","Url")
'		If Environment("ApplicationPath") = "STQA" Then
'			strURL = "https://qa01ebc.thehartford.com"
'			wait 2
'			Browser("EBC_Login").Page("Login").WebEdit("USERText").Set "NBVQB142"
'			Browser("EBC_Login").Page("Login").WebEdit("PASSWORD").Set "password"
'			Browser("EBC_Login").Page("Login").Image("Login").Click
'			wait 1
'			Browser("EBC_Login").Page("Login").Link("Quote Small Commercial").Click
'			Browser("EBC_Login").Close
'			Exit Sub
'	
'	
'			
'		ElseIf Environment("ApplicationPath") = "LTI" Then
'			strURL = "https://ti04esubmissions.thehartford.com/login/icon1Login.fcc?TYPE=33554433&REALMOID=06-91ecc112-de08-4f4d-9ce9-ba4a299ed23b&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=SrAjRsDtPnvPdTovOaoYONasJSqdRFGiJ2zxvP6oOrfDma045m1Uuu3HEDs3Lf02&TARGET=-SM-HTTPS%3a%2f%2fti04esubmissions%2ethehartford%2ecom%2fcustmgmt%2fam%2ftab%2exhtml%3ftabId%3dvi98onahgpnub1vq6tljnqma5r"
'		Else
'	'		strURL = "https://ti04esubmissions.thehartford.com/login/icon1Login.fcc?TYPE=33554433&REALMOID=06-91ecc112-de08-4f4d-9ce9-ba4a299ed23b&GUID=&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=SrAjRsDtPnvPdTovOaoYONasJSqdRFGiJ2zxvP6oOrfDma045m1Uuu3HEDs3Lf02&TARGET=-SM-HTTPS%3a%2f%2fti04esubmissions%2ethehartford%2ecom%2fcustmgmt%2fam%2ftab%2exhtml%3ftabId%3dvi98onahgpnub1vq6tljnqma5r"
'         	strURL = "https://qa04esubmissions.thehartford.com/custmgmt/am_opener.jsp?src=ebc"
'	'		strURL = "https://qa01ebc.thehartford.com"
'		End If
'	'	Systemutil.Run "iexplore", "https://qa04esubmissions.thehartford.com/custmgmt/am_opener.jsp?src=ebc" ' LTQA
		Systemutil.Run "iexplore",strURL
		wait 2
		With Browser("ICONSmallCommercial").Page("ICON")
			If .Link("LNK_LogOff").Exist(1) Then
				Exit Sub
				.Link("LNK_LogOff").Click
				
	'			If Browser("Logout HTML Page 09142010-1600").Dialog("Windows Internet Explorer").Exist(1) Then
	'				Browser("Logout HTML Page 09142010-1600").Dialog("Windows Internet Explorer").WinButton("Yes").Click
	'			End If
				'Wait_Obj_Existence(Browser("ICONSmallCommercial").Page("ICON").WebEdit("EDI_USER"))
				Systemutil.Run "iexplore",strURL
			End If
		End With
	'	Systemutil.Run "iexplore", Environment("EnvURL")
	'		if Browser("ICONSmallCommercial").Link("Log Off").Exist(1) Then
	'		
	'			Browser("ICONSmallCommercial").Link("Log Off").Click
			'Systemutil.Run "iexplore", "https://qa01ebc.thehartford.com"
	'			Browser("ICONSmallCommercial").WaitProperty "height",">200",Environment("Wait_time")
	'			call Userlogin()	
	'		Else	
				call Userlogin(strURL)	
				
	'		End if
	Else
		Browser("ICONSmallCommercial").Page("NewCustomer").WebElement("ELE_Home").Click
		If Browser("ICONSmallCommercial").Page("NewCustomer").Link("LNK_Continue").Exist(1) then
			Browser("ICONSmallCommercial").Page("NewCustomer").Link("LNK_Continue").click
		End If
		

	End If
	If Browser("ICONSmallCommercial").Page("NewCustomer").Link("Lnk_Close").Exist(1) Then
		Browser("ICONSmallCommercial").Page("NewCustomer").Link("Lnk_Close").Click
	End If
End Sub


'###################################################################################################################
	
'###################################################################################################################
'Function Description   : Function to fill the Username and password in Login page
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 19/09/2015
'###################################################################################################################
Sub Userlogin(strURL)
''	intflag = 0
''	Do
''		With Browser("ICONSmallCommercial").Page("ICON")
''		If .Link("LNK_LogOff").Exist(1) Then
''			.Link("LNK_LogOff").Click
''			wait 3
''			If Browser("Logout HTML Page 09142010-1600").Dialog("Windows Internet Explorer").Exist(1) Then
''				Browser("Logout HTML Page 09142010-1600").Dialog("Windows Internet Explorer").WinButton("Yes").Click
''			End If
'			'Wait_Obj_Existence(Browser("ICONSmallCommercial").Page("ICON").WebEdit("EDI_USER"))
'			Systemutil.Run "iexplore",strURL
'		End If
		wait 2
	With Browser("ICONSmallCommercial").Page("ICON")
		If .WebEdit("EDI_USER").Exist(1) Then
			Nbv_user=gobjDataTable.GetData("General_Data","EDI_USER")
			Nbv_pwd=gobjDataTable.GetData("General_Data","EDI_PASSWORD")
'			Nbv_user=gobjDataTable.GetData("Test_Data","EDI_USER")
'			Nbv_pwd=gobjDataTable.GetData("Test_Data","EDI_PASSWORD")
			.WebEdit("EDI_USER").Set Nbv_user
			.WebEdit("EDI_PASSWORD").Set Nbv_pwd
			.WebButton("BTN_Login").Click
'			Wait_Obj_Existence(.Link("LNK_AddNewCustomer"))
			Exit Sub
		End If
'		Wait 1	
'		intflag	= intflag + 1	
	End With	
'	Loop While intflag < Environment("Wait_time")
End Sub

'###################################################################################################################
	
'###################################################################################################################
'Function Description   : Function to output the expected QuoteID
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 21/09/2015
'###################################################################################################################
Sub Extract_QuoteID()

'wait 3
	If Browser("ICON_NBV").Page("NBV").WebElement("ELE_PremiumSummary").Exist(20) Then
		QuoteID = Browser("ICON_NBV").Page("NBV").WebElement("ELE_QuoteID").GetROProperty("innertext")
	'	Browser("ICON_NBV").Close
		QuoteID=mid(QuoteID,9)
		strPol_ID = Right(QuoteID,6)
		gobjReport.UpdateTestLog "Quote ID","The Quote ID is" & QuoteID, "Pass"
		End If
    If Browser("ICON_NBV").Page("PremiumSummary").WebElement("General Information").Exist(1) Then
        QuoteID = Browser("ICON_NBV").Page("NBV").WebElement("ELE_QuoteID").GetROProperty("innertext")
		QuoteID=mid(QuoteID,9)
		strPol_ID = Right(QuoteID,6)
		gobjReport.UpdateTestLog "Quote ID","The Quote ID is" & QuoteID, "Pass"

	Else
		gobjReport.UpdateTestLog "Premium Summary","Quote is not moved to Premium summary screen", "Fail"
	End If
'	If Browser("ICONSmallCommercial").Page("ICON").Link("LNK_LogOff").Exist(1) Then 
'		Browser("ICONSmallCommercial").Page("ICON").Link("LNK_LogOff").Click
'	End If
	Call OutputQuoteID("Edi_PolNumber",strPol_ID)
'	Call Outputdata("Policy_num",strPol_ID)
End Sub
'###################################################################################################################
'Function Description   : Function to output the expected PolicyID
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 21/09/2015
'###################################################################################################################
Sub Extract_PolicyID()
	strPolicyID = Browser("ICON_NBV").Page("NBV").WebElement("ELE_QuoteID").GetROProperty("innertext")
	strPol_ID = Right(strPolicyID,6)
	'Msgbox "The Policy ID is " & strPolicyID 
	Call Outputdata("Edi_PolNumber",strPol_ID)
'	Call Outputdata("POL_SYS_ADDR",strPolicyID)
End Sub


'###################################################################################################################
'Function Description   : Function to fill the Answers for underwriter Questions
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 06/10/2015
'###################################################################################################################
Function UnderWriting_Ques()
	'UWResp = gobjDataTable.GetData(Environment("strCurrentKeyword"), "CAL_UnderWriting_Ques")
	'UWResp = gobjDataTable.GetData("Test_Data", "Underwriter_Response")
'	If UWResp <> "" Then
'		Call Fill_UnderWriterQues(UWResp)
'	Else
	wait 3
		Set ODesc = Description.Create
		ODesc("micclass").value = "WebRadioGroup"
		Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
		arrobjCnt = arrobj.Count
		If arrobjCnt>0 Then
			For inc = 0 To arrobjCnt-1
				arrobj(inc).Select "NO"	
				arrobjCnt = UpdateCount("WebRadioGroup",arrobj.Count)
				Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
				Set HardStop = Browser("ICON_NBV").Page("NBV").WebElement("class:=hardStop")
				If HardStop.Exist(3) Then
					arrobj(inc).Select "YES"
					arrobjCnt = UpdateCount("WebRadioGroup",arrobj.Count)
					Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
				End If
			Next
			
		End If
		
		wait(1)
		Set ODesc = Description.Create
		ODesc("micclass").value = "WebList"
		ODesc("visible").value = True
		Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
		arrobjCnt = arrobj.Count
		If arrobjCnt>0 Then
		For inc = 0 To arrobj.Count-1
			arrobj(inc).Select(1)
			Next
		End If
		
		wait(1)
		
		Set ODesc = Description.Create
		ODesc("micclass").value = "WebCheckBox"
		Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
		'arrobjCnt = 0
		arrobjCnt = UpdateCount("WebCheckBox",arrobj.Count)
		If arrobj.Count>0 Then
			For inc = 0 To arrobj.Count-1
				arrobj(inc).set "ON"
				Set HardStop = Browser("ICON_NBV").Page("NBV").WebElement("class:=hardStop")
				If HardStop.Exist(0) Then
					arrobj(inc).set "OFF"
				Else
				 	Exit For
				End If
			Next
			
		End If
		wait(1)
		Set ODesc = Description.Create
		ODesc("micclass").value = "WebEdit"
		ODesc("type").value = "text.*"
		Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
		'arrobjCnt = UpdateCount("WebEdit",arrobj.Count)
		If arrobj.count>0 Then
			For inc = 0 To arrobj.count-1
			If arrobj(inc).GetROProperty("max length") <= 3 Then
				arrobj(inc).set "2"	
				Else 
				arrobj(inc).set "Test"
			End If
		Next
			
		End If
		
		Set ODesc = Description.Create
		ODesc("micclass").value = "WebRadioGroup"
		ODesc("checked").value = 0
		Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
		arrobjCnt = arrobj.Count
		If arrobjCnt>0 Then
			Call UnderWriting_Ques()
			Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
			
			Else 
			Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
		End If
		
		'Wait(2)
'	End If
'	Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
	
'	 Environment("CurrentTime_1") = time()  
' If  Browser("ICON_NBV").Page("NBV").WebElement("ELE_PremiumSummary").Exist(40) then
'     Environment("CurrentTime_2") = time()  
'     Seconds = Datediff("s",Environment("CurrentTime_1"),Environment("CurrentTime_2"))
' End if
 	'QCUtil.CurrentRun.Field("RN_DRAFT") = "N"
End Function

'##################################################################################################################################
'Function Description   : Function to update the count of Fields(Eg- Radiobutton,Editbox,checkbox) for UnderWriting_Ques function 
'Input Parameters		: ObjType,Count
'Author					: Priyanka
'Date Created			: 06/10/2015
'##################################################################################################################################
Function UpdateCount(ObjType,Count)
	Set ODesc = Description.Create
	ODesc("micclass").value = ObjType
	Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
	UpdateCount = arrobj.Count
End Function

Function ExecuteQuery(strQuery)
Dim strFilePath, strConnectionString, r_objConn,Cmd
        strFilePath = "C:\Users\COG11704\Desktop\IAS1\Run Manager.xls"
            strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConnect = CreateObject("ADODB.Connection")
        r_objConnect.Open strConnectionString
		
		Dim strNonQuery, objTestData
		strTestDataSheet = "Common_Data"


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


'##################################################################################################################################
'Function Description   : Function to Fill Class Description and amount payroll value 
'Input Parameters		: None
'Author					: Ankit
'Date Created			: 12/10/2015
'##################################################################################################################################
Function Fill_ClassDescription()
	Set ODescLst = Description.create
	Set ODescEdi = Description.create
	ODescLst("micclass").value ="Weblist"
	ODescLst("name").value ="ClassDescription.*"
	ODescEdi("micclass").value ="WebEdit"
	ODescEdi("name").value ="AnnualClassPayroll.*"
	Set arrobjLst = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODescLst)
	Set arrobjEdi = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODescEdi)
	LstCnt = arrobjLst.count
	If LstCnt>0 Then
		For inc = 0 To LstCnt-2
			arrobjLst(inc).select(1)
			arrobjEdi(inc).Set gobjDataTable.GetData("StateLocationClasscode", "EDI_Payroll")
			
	     Next
	End If
	Set ODescLst = Nothing
	Set ODescEdi = Nothing
Wait(2)
End Function




Function AddLocation(strData)
Obj_Browser = gobjDataTable.GetData(Environment("strCurrentKeyword"), "Browser")
 Obj_Page = Environment("strCurrentKeyword")
 Set PrntObj = Browser(Obj_Browser).Page(Obj_Page)
 strDataLoc = Replace(strData,".*","")
arrAllItems = Split(PrntObj.WebList("LST_Location_CC").GetROProperty("all items"), ";")
 For inc = 0 To Ubound(arrAllItems)
 	If Instr(1,arrAllItems(inc),strDataLoc) Then
 		PrntObj.WebList("LST_Location_CC").Select(arrAllItems(inc))
 		Exit For
 	End If
 Next
 End Function

Function ValidateListItems(strData)
	If strData(inc)= "random" Then
		incwait=0
		objcount=Browser("ICON_NBV").Page("ClassCodeschedule").WebList("LST_ClassDescription_CC").count
Do
	
	If objcount<1 Then
	wait 1
	   incwait = incwait+1	
		
	Else
		Exit function
	End If
	Loop While (incwait < 3)
	gobjReport.UpdateTestLog "Issue Policy","The listitem was not fetched successfully", "Fail"
	'MsgBox "The Object does not Exist in current Screen." & vbCrLf & "Make sure you are navigating on correct screen", ,"Object Does Not EXIST"

	End If
End Function
'Function GetDataDetails(Obj,strData)
''	strWriteData = strData
''	arrData = Split(strData,",")
''	If Ubound(arrData) <> 0 Then
''		strData = arrData(0)
''	End If
'
'			If strData = "Random" Then
'				strData = Random_Data(Left(Obj,3))
'		   End If
''			If Instr(1,strData,"*") <> 0 Then
''				strWriteData = Replace(strData,"*","")
''				strData = strWriteData
''				If Obj = strWriteData Then
''					strWriteData = OutputData(strWriteData)
''					strData = "Skip"
''				End If
''				Call OutputCommondata(Obj,strWriteData)
''				End If 
''				arrObj = "Skip"
''			ElseIf Instr(1,Obj,"$") <> 0 Then
''				'Call VerifyData(arrObj,Replace(strData,"$",""))
''				Call VerifyData(Replace(Obj,"$",""))  
''				arrObj = "Skip"
'			If Instr(1,strData,"#") <> 0 Then
'				strData =  Getcommonoutdata(Replace(strData,"#",""))
'			End If
'
'			GetDataDetails = strData
'End Function
'



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



'###################################################################################################################
'Function Description   : Function to Answer Eligibility Questions in GeneralInformation page
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 15/11/2015
'###################################################################################################################
Function Eligibility_Ques()
Wait 9
		Environment("strCurrentKeyword") = "GeneralInformation"
		If Browser("ICON_NBV").Page("GeneralInformation").WebElement("ELE_GeneralInformation").Exist(5) Then
			Set ODesc = Description.Create
			ODesc("micclass").value = "WebRadioGroup"
			Set arrobj = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODesc)
			arrobjCnt = arrobj.Count
			If arrobjCnt>0 Then
				For inc = 0 To arrobjCnt-1
				arrobj(inc).Select "NO"
                arrobjCnt = UpdateCount("WebRadioGroup",arrobj.Count)				
				Set HardStop = Browser("ICON_NBV").Page("NBV").WebElement("class:=hardStop")
				If HardStop.Exist(3) Then
					arrobj(inc).Select "YES"
					arrobjCnt = UpdateCount("WebRadioGroup",arrobj.Count)
					Set arrobj = Browser("ICON_NBV").Page("NBV").ChildObjects(ODesc)
				End If
				wait(3)
				Next
			End If
		
			Set ODesc = Description.Create
		    ODesc("micclass").value = "WebCheckBox"
'			ODesc("innertext").value="None of the Above"
'			ODesc("innerhtml").value="None of the Above"
			ODesc("visible").value = False
			Set arrobj = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODesc)
			'MsgBox arrobj.count
			If arrobj.Count>0 Then
				For inc = 0 To arrobj.Count-1
				arrobj(arrobj.Count-1).set "ON"
				Exit For
				Next
			End If
End If
End Function

'###################################################################################################################
'Function Description   : Function to set fields value with current browser and page.
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 27/10/2015
'###################################################################################################################
Function setObjType(ByRef strObj)
    If (Instr(1,UCASE(Environment("strCurrentKeyword")),"VEHTYPES")) <> 0 Then
    	Environment("strCurrentKeyword") = "VEHTYPES_Add"
    Else 
    Obj_Browser = gobjDataTable.GetData(Environment("strCurrentKeyword"), "Browser")
    Obj_Page = Environment("strCurrentKeyword")
    End IF 
	If Obj_Browser="OBSC" Then
	 Obj_Frame = "iframe" 
		strType = UCASE(Split(strObj,"_")(0))
	    Set PrntObj = Browser(Obj_Browser).Page(Obj_Page).Frame(Obj_Frame)
	    Select Case strType
	        Case "EDI"
	            set CrntObj = PrntObj.WebEdit(strObj)
	        Case "LNK"
	            set CrntObj = PrntObj.Link(strObj)
	            strObj = "CLK_"&strObj
	        Case "ELE"
	            set CrntObj = PrntObj.WebElement(strObj)
	            strObj = "CLK_"&strObj
	        Case "BTN"  
	            set CrntObj = PrntObj.WebButton(strObj)
	            strObj = "CLK_"&strObj
	        Case "IMG"  
	            set CrntObj = PrntObj.Image(strObj)
	            strObj = "CLK_"&strObj
	        Case "RDO"
	            set CrntObj = PrntObj.WebRadioGroup(strObj)
	        Case "CHK"
	            set CrntObj = PrntObj.WebCheckbox(strObj)
	        Case "LST"
	            set CrntObj = PrntObj.WebList(strObj)
            
	        Case Else
	        	set CrntObj = PrntObj
       End Select
       Set setObjType = CrntObj
	Else
	    strType = UCASE(Split(strObj,"_")(0))
	    If strType <> "MFO" Then
	    	Set PrntObj = Browser(Obj_Browser).Page(Obj_Page)
	    Else
	    	Set PrntObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen")  
	    End If
	    	    
	    
	    Select Case strType
	        Case "EDI"
	            set CrntObj = PrntObj.WebEdit(strObj)
	        Case "LNK"
	            set CrntObj = PrntObj.Link(strObj)
	            strObj = "CLK_"&strObj
	        Case "ELE"
	            set CrntObj = PrntObj.WebElement(strObj)
	            strObj = "CLK_"&strObj
	        Case "BTN"  
	            set CrntObj = PrntObj.WebButton(strObj)
	            strObj = "CLK_"&strObj
	        Case "IMG"  
	            set CrntObj = PrntObj.Image(strObj)
	            strObj = "CLK_"&strObj
	        Case "RDO"
	            set CrntObj = PrntObj.WebRadioGroup(strObj)
	        Case "CHK"
	            set CrntObj = PrntObj.WebCheckbox(strObj)
	        Case "LST"
	            set CrntObj = PrntObj.WebList(strObj)
            Case "FLE"
                set CrntObj = PrntObj.WebFile(strObj)
            Case "SRH"
                set CrntObj = PrntObj.WebElement(strObj)
	        Case "MFO"
'	        	strObj = Replace("MFO_","")
'	        	strQuery = "Select * from [CLA_Repository$] where Obj_Name = '" & strObj & "'"
'				Set RSObjID = ExtractSceKeywords (strQuery)
				strID = Split(strObj,"_")(1)
				If IsNumeric(strID) Then
					Set CrntObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:="&strID)					
				Else
					If strID <> "" Then
						arrstrID = Split(strID,",")
						If Ubound(arrstrID) <> 0 Then
							intindex = arrstrID(1)
						Else
							intindex = 0
						End If
					Set CrntObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("attached text:="&arrstrID(0),"Index:="&intindex)
					
					Else
						Set CrntObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen")					
					End If
					
				End If
				'Set CrntObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:="&strID)
				Set RSObjID = Nothing
	        Case Else
	        	set CrntObj = PrntObj
       End Select
       
		If Environment("strCurrentKeyword") = "StateratingPlanoption" Then
			Call SetStateDefaultProp()
			str_State = gobjDataTable.GetData(Environment("strCurrentKeyword"), "State")
			strhtmlid = CrntObj.GetTOProperty("html id")
			If strhtmlid <> "" Then
				strhtmlid = Replace(strhtmlid,".*",str_State)
				CrntObj.SetTOProperty "html id", strhtmlid
			End If
			If CrntObj.Exist(1) <> True Then
				strObj = "Skip"
			End If
       End If
       
		If Environment("strCurrentKeyword") = "PricingAdjustment" Then
			Call SetPricingDefaultProp()
			str_State = gobjDataTable.GetData(Environment("strCurrentKeyword"), "State")
			strhtmlid = CrntObj.GetTOProperty("html id")
			If strhtmlid <> "" Then
				strhtmlid = Replace(strhtmlid,".*",str_State)
				CrntObj.SetTOProperty "html id", strhtmlid
			End If
			If CrntObj.Exist(1) <> True Then
				strObj = "Skip"
			End If
       End If
     Set setObjType = CrntObj  
 End If 
 
End Function


'###################################################################################################################
'Function Description   : Function to set default property with respect to the state in stateRatePlanOption Page
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 02/12/2015
'###################################################################################################################
Function SetStateDefaultProp()
	With Browser("ICON_NBV").Page("StateratingPlanoption")
		.WebCheckBox("CHK_DrugFreeCredit").SetTOProperty "html id", "DrugFree.*"
		.WebCheckBox("CHK_ExtBroadFrom").SetTOProperty "html id", "ExtendBroadForm.*"
		.WebCheckBox("CHK_MedicalCredit").SetTOProperty "html id", ".*DesignatedMedicalProvider"
		.WebCheckBox("CHK_SafetyCredit").SetTOProperty "html id", "SafetyCredit.*"
		.WebEdit("EDI_ARAP").SetTOProperty "html id", "Arap.*"
		.WebEdit("EDI_ExpMode").SetTOProperty "html id", "ExperMod.*"
		.WebList("LST_Waiver").SetTOProperty "html id", "WaiverOfSubrogation.*"
		.WebList("LST_CCPA").SetTOProperty "html id", "ContractorsClassPremium.*"
		.WebCheckBox("CHK_MedicalCredit").SetTOProperty "html id", ".*ManagedCareCredit"
		.WebList("LST_CRPM").SetTOProperty "html id", ".*RiskManagement"
		.WebCheckBox("CHK_Deliberateintent").SetTOProperty "html id", ".*DeliberateIntent"
	End With
End Function


'###################################################################################################################
'Function Description   : Function to set default property with respect to the state in PricingAdjustment Page
'Input Parameters		: None
'Author					: Shankar
'Date Created			: 02/12/2015
'###################################################################################################################
Function SetPricingDefaultProp()
	With Browser("ICON_NBV").Page("PricingAdjustment")
		.WebList("LST_FlexRateAdjustment").SetTOProperty "html id", "FlexRateAdjustment.*"
		.WebList("LST_PricingAdjustment").SetTOProperty "html id", "PricingAdjustmentFactor.*"
		.WebList("LST_ScheduleMod").SetTOProperty "html id", "ScheduleModification.*"
	End With
End Function

'###################################################################################################################
'Function Description   : Function to fill producercode in New Customer page.
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 08/11/2015
'###################################################################################################################
Function NewCustomer_ProducerCode()
     strData = gobjDataTable.GetData(Environment("strCurrentKeyword"), "CAL_NewCustomer_ProducerCode")
      If Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebEdit("EDI_ProducerCode2").Exist(1) Then
		strProducerCode = Left(strData,8)	 
		Set myDeviceReplay = CreateObject("Mercury.DeviceReplay")
		Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebEdit("EDI_ProducerCode2").Click
		myDeviceReplay.SendString strProducerCode
		myDeviceReplay.presskey 15
		Wait 1
		'Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebEdit("EDI_ProducerCode2").Set strProducerCode
	   ElseIf Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebList("LST_ProducerCode").Exist(1) Then	   		
		   'Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebList("LST_ProducerCode").Select strData
		   allproducercodes = Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebList("LST_ProducerCode").GetROProperty("all items")
		   allproducercodes = split(allproducercodes, ";" )		   
		   For inc = 0 To Ubound(allproducercodes) Step 1
		   	if instr(1,allproducercodes(inc), strData) then
		   	strData = allproducercodes(inc)
		   	Exit for
		   	End If
		   Next
	   		Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebList("LST_ProducerCode").Select strData
      End If
End Function

'###################################################################################################################
'Function Description   : Function to Click Edit and update the required fields in Classcode Schedule page
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 15/12/2015
'###################################################################################################################
Function EditClass()
strData = gobjDataTable.GetData("ClassCodeschedule", "Delete")
If strData <> "" Then
    If Browser("ICON_NBV").Page("ClassCodeschedule").WebTable("ClasscodeDelete").Exist(1) Then
		RwCnt = Browser("ICON_NBV").Page("ClassCodeschedule").WebTable("ClasscodeDelete").GetROProperty("rows")
		For i = 2 To RwCnt
		Browser("ICON_NBV").Page("ClassCodeschedule").WebTable("ClasscodeDelete").ChildItem(2,6,"Link",1).Click
		Browser("ICON_NBV").Page("ClassCodeschedule").Link("Yes").Click
		Next
	End If 
Else
	Set ODesc = Description.Create
	Environment("strCurrentKeyword")="ClassCodeschedule"
	ODesc("micclass").value = "Link"
	ODesc("name").value = "Edit"
	ODesc("visible").value = True
	Set arrobj = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODesc)
	arrobjCnt = arrobj.Count
	If arrobjCnt>0 Then
		For inc = 0 To arrobjCnt-1
        wait 3
		Set arrobj = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODesc)
	
		arrobj(inc).Click
		Set strrdocheck=Browser("ICON_NBV").Page("ClassCodeschedule").WebRadioGroup("RDO_OfficerCC")
		wait 1
		If strrdocheck.Exist(1) Then
         	strrdocheck.Select "0"
        End If
            Browser("ICON_NBV").Page("ClassCodeschedule").WebList("LST_ClassDescription_CC").highlight
			Browser("ICON_NBV").Page("ClassCodeschedule").WebList("LST_ClassDescription_CC").Select(1)
			strPayroll = gobjDataTable.GetData(Environment("strCurrentKeyword"), "AnnualPayroll")
		If strPayroll <> "" Then
		Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).WebEdit("EDI_AnnualPayroll").Set strPayroll
		else
		Browser("ICON_NBV").Page("ClassCodeschedule").WebCheckBox("CHK_IfAny").Set "ON"
		End If
        Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).WebButton("BTN_Update").Click
        wait 10
        Next
   End If
End If
		
End Function

'###################################################################################################################
'Function Description   : Function to verify the page name in Pricing Justification screen
'Input Parameters		: None
'Return Value    		: call Pricingjustification() or call PricingFactorjustification()
'Author					: Priyanka
'Date Created			: 27/12/2015
'###################################################################################################################
Function PageCheck()
 	If Browser("ICON_NBV").Page("PricingJustification").WebElement("Percentage").Exist Then
 	Call Pricingjustification()
 	ElseIf Browser("ICON_NBV").Page("PricingJustification").WebElement("PricingFactor").Exist Then
 	Call PricingFactorjustification()
 	End If
   	
End Function

'###################################################################################################################
'Function Description   : Function to fill fields in pricing justification page
'Input Parameters		: None
'Return Value    		: call PageCheck()
'Author					: Priyanka
'Date Created			: 27/12/2015
'###################################################################################################################
Function Pricingjustification()
	
	If Browser("ICON_NBV").Page("PricingJustification").WebElement("Percentage").Exist Then
      PerValue = Browser("ICON_NBV").Page("PricingJustification").WebElement("Percentage").GetROProperty("innertext")
    If PerValue > +5 OR PerValue < -5 Then
		Set Odesc = Description.Create
		ODesc("micclass").value = "WebList"
		ODesc("name").Value = ".*RiskPercentage.*"
		ODesc("visible").value = True
		Set arrobj = Browser("ICON_NBV").Page("PricingJustification").ChildObjects(ODesc)
		arrobjCnt = arrobj.Count
		PerNum = Replace(PerValue,Right(PerValue,1), "")
		ObjValue = PerNum/(arrobjCnt-1)
		Lstvalue = Split(ObjValue,".") 
		ModValue = (PerNum MOD (arrobjCnt-1)) & "%"
		'row1 = ((PerNum MOD arrobjCnt) + Lstvalue(0)) & "%"
		arrobj(0).Select ModValue
		If Lstvalue(0) = "-0" Then
		   For inc = 1 To arrobjCnt-1
		arrobj(inc).Select "0%"
		Next
			Set ODesc = Description.Create
			    ODesc("micclass").value = "WebList"
			    ODesc("visible").value = True
			    ODesc("name").Value = ".*RiskExplanation.*"
				Set arrobjL = Browser("ICON_NBV").Page("PricingJustification").ChildObjects(ODesc)
				arrobjCntLST = arrobjL.Count
		        If arrobjCntLST>0 Then
		           arrobjL(0).select(1)
		           For inc = 1 To arrobjCnt-1
					arrobjL(inc).select(0)
				   Next
		        End If 
		Else
		For inc = 1 To arrobjCnt-1
		arrobj(inc).Select Lstvalue(0) & "%"
		Next
		
		Set ODesc = Description.Create
			    ODesc("micclass").value = "WebList"
			    ODesc("visible").value = True
			    ODesc("name").Value = ".*RiskExplanation.*"
				Set arrobjL = Browser("ICON_NBV").Page("PricingJustification").ChildObjects(ODesc)
				arrobjCntLST = arrobjL.Count
		        If arrobjCntLST>0 Then
		           For j = 0 To arrobjCntLST-1
		               arrobjL(j).select(1)
		              Next
		         End If
	   End If
	   Browser("ICON_NBV").Page("PricingJustification").WebButton("BTN_Continue").Click
	   Call PageCheck()
	End If
End If
End Function


'#####################################################################################################################################
'Function Description   : Function to fill fields in pricing justification page only if the page contains PricingFactor related fields. 
'Input Parameters		: None
'Return Value    		: call PageCheck()
'Author					: Priyanka
'Date Created			: 27/12/2015
'#####################################################################################################################################
Function PricingFactorjustification()

If Browser("ICON_NBV").Page("PricingJustification").WebElement("PricingFactor").Exist Then
	Set ODesc = Description.Create
	ODesc("micclass").value = "WebList"
	ODesc("visible").value = True
	ODesc("name").Value = ".*RiskExplanation.*"
	Set arrobjL = Browser("ICON_NBV").Page("PricingJustification").ChildObjects(ODesc)
	arrobjCntLST = arrobjL.Count
	If arrobjCntLST>0 Then
		For j = 0 To arrobjCntLST-1
		arrobjL(j).select(1)
		Next
	End If
	Set ODesc = Description.Create
	ODesc("micclass").value = "WebCheckBox"
	ODesc("name").Value = ".*RiskAttribute.*"
	Set arrobjR = Browser("ICON_NBV").Page("PricingJustification").ChildObjects(ODesc)
	arrobjCntR = arrobjR.Count
	If arrobjCntR>0 Then
		For k = 0 To arrobjCntR-1
		arrobjR(k).set "ON"
		Next
		Browser("ICON_NBV").Page("PricingJustification").WebButton("BTN_Continue").Click
		Call PageCheck()
	End If
End If 	
End Function

'#####################################################################################################################################
'Function Description   : Function to Click Edit and update the required fields in Location Address Information page 
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 02/01/2016
'#####################################################################################################################################
Function LocAddInfo()
RowCnt = Browser("ICON_NBV").Page("LocationAddressInfo").WebTable("State").GetRoProperty ("rows")
For i = 2 To RowCnt
a = Browser("ICON_NBV").Page("LocationAddressInfo").WebTable("State").GetCellData(i,7)
If instr(a,"Required") <> 0 Then
	Browser("ICON_NBV").Page("LocationAddressInfo").WebTable("State").ChildItem(i,8,"Link",0).Click
	Browser("ICON_NBV").Page("LocationAddressInfo").WebEdit("Edi_Phone").Set "(891) 787-3413"
	Browser("ICON_NBV").Page("LocationAddressInfo").WebButton("BTN_Update").Click
End If
Next
Browser("ICON_NBV").Page("LocationAddressInfo").WebButton("BTN_Continue").Click
End Function


'#####################################################################################################################################
'Function Description   : Function to fill customers Bank and pament Information in Payment Details page 
'Input Parameters		: None
'Author					: Priyanka
'Date Created			: 05/01/2016
'#####################################################################################################################################
Function PaymentDetails()

	With Browser("ICON_NBV").Page("BillingInfo").Frame("PaymentDetails")
	If .Exist Then
    Environment("strCurrentKeyword") = "PaymentDetails"
    PaymentType = gobjDataTable.GetData(Environment("strCurrentKeyword"), "Payment_Type")
	.Link(PaymentType).Click
    RoutingNum = gobjDataTable.GetData(Environment("strCurrentKeyword"), "RoutingNum")
    .WebEdit("EDI_RoutingNumber").Set RoutingNum
    BankAccNum = gobjDataTable.GetData(Environment("strCurrentKeyword"), "BankAccNum")
    .WebEdit("EDI_acctNumber").Set BankAccNum
    OtherAmt = gobjDataTable.GetData(Environment("strCurrentKeyword"), "OtherAmt")
    PaymentAmtRDO = gobjDataTable.GetData(Environment("strCurrentKeyword"), "PaymentAmt")
    .WebRadioGroup("RDO_CurrentBal").Select PaymentAmtRDO
    If PaymentAmtRDO = "#2" Then
       .WebEdit("EDI_OtherAmount").Set OtherAmt
    End If
    .WebButton("BTN_Next").Click
    .WebButton("BTN_Submit").Click
    .WebButton("BTN_Finish").Click
End If
End With

End Function

Function LoginECOS
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
     strURL = "http://sit2-ecos.thehartford.com/cc/ClaimCenter.do"
  'Close All Browsers
   SystemUtil.CloseProcessByName "iexplore.exe"
   SystemUtil.CloseProcessByName "chrome.exe"
   Systemutil.Run "iexplore",strURL
		call UserECOSlogin()
End Function

Sub UserECOSlogin()
''	intflag = 0
''	Do
''		With Browser("ICONSmallCommercial").Page("ICON")
''		If .Link("LNK_LogOff").Exist(1) Then
''			.Link("LNK_LogOff").Click
''			wait 3
''			If Browser("Logout HTML Page 09142010-1600").Dialog("Windows Internet Explorer").Exist(1) Then
''				Browser("Logout HTML Page 09142010-1600").Dialog("Windows Internet Explorer").WinButton("Yes").Click
''			End If
'			'Wait_Obj_Existence(Browser("ICONSmallCommercial").Page("ICON").WebEdit("EDI_USER"))
'			Systemutil.Run "iexplore",strURL
'		End If
		wait 2
	With Browser("ECOSBrowser").Page("Login")
	'Browser("ECOSBrowser").Frame("PAGEECOS").CaptureBitmap

		If .WebEdit("EDI_NetworkID").Exist(1) Then
			Ecos_user=gobjDataTable.GetData("General_Data","EDI_NetworkID")
			Ecos_pwd=gobjDataTable.GetData("General_Data","EDI_PASSWORD")
'			Nbv_user=gobjDataTable.GetData("Test_Data","EDI_USER")
'			Nbv_pwd=gobjDataTable.GetData("Test_Data","EDI_PASSWORD")
			.WebEdit("EDI_NetworkID").Set Ecos_user
			.WebEdit("EDI_PASSWORD").Set Ecos_pwd
			.WebElement("ELE_Log In").Click
'			Wait_Obj_Existence(.Link("LNK_AddNewCustomer"))
			Exit Sub
		End If
'		Wait 1	
'		intflag	= intflag + 1	
	End With	
'	Loop While intflag < Environment("Wait_time")
End Sub


Function CreateXML()

Browser("AgencyPortalXML").Page("AgencyPortal Console").WebList("select").Select "Current AP simplified ACORD XML"
XMLData1= Browser("AgencyPortalXML").Page("AgencyPortal Console").WebXML("XMLData").GetROProperty("innertext")
set filesys=CreateObject("Scripting.FileSystemObject") 
set Writefile = filesys.CreateTextFile("H:\XML\" & gobjTestParameters.CurrentTestcase &".txt")
Writefile.Write(XMLData1)

'  Call VerifyXML()
End Function

Function VerifyXML() 
'Taking Node value in xml 2
   Set xmldoc = CreateObject("Microsoft.XMLDOM")
   path = "H:\XMLData2.xml"
   xmldoc.load(path)
   Set Ele2 = xmldoc.documentElement.childnodes
   'VerifyName = "BOPPolicyQuoteInqRq/RqUID"
   VerifyName = "WorkCompPolicyQuoteInqRq/RqUID"
   Set Verifyval = xmldoc.SelectSingleNode(VerifyName) 
   'NodeNam = Verifyval.nodename
   NodeVal2 = Verifyval.text
   
   
'Taking Node value in xml 1
                path = "H:\XMLData1.xml"
                xmldoc.load(path)
                Set Ele1 = xmldoc.documentElement.childnodes
                'VerifyName = "BOPPolicyQuoteInqRq/RqUID"
                VerifyName = "WorkCompPolicyQuoteInqRq/RqUID"
                Set Verifyval = xmldoc.SelectSingleNode(VerifyName)
                'NodeNam = Verifyval.nodename
                NodeVal1 = Verifyval.text

'comparing both xmls

                If instr(1,NodeVal1,NodeVal2)<>0 Then
                                msgbox "Pass"
                                msgbox NodeVal1
                                msgbox NodeVal2
                End If


End Function

Function VerifyObject(arrVerify)

	
	Obj_Browser = gobjDataTable.GetData(Environment("strCurrentKeyword"), "Browser")
     Obj_Page = Environment("strCurrentKeyword")

		If   Obj_Browser <> ""  Then
			Set PrntObj = Browser(Obj_Browser).Page(Obj_Page)
		End If     
      Dim incwait
        incwait = 0 	
'	For inc = 1 to Ubound(arrVerify)	
	
		If Instr(1,arrVerify(1),"%") <> 0 Then
			Call VerificationFields(arrVerify)
		End If
				If Instr(1,arrVerify(1),"*") <> 0 Then
					ObjectName = Replace(arrVerify(1),"*","")
					ObjectType = Mid(ObjectName,1,3)
					
					Select Case ObjectType 
						 Case "EDI"
							 ObjectValue = PrntObj.WebEdit(ObjectName).GetROProperty("value")
							 Call OutputCommondata(ObjectName,ObjectValue)
						 Case "LNK"
							 ObjectValue = PrntObj.Link(ObjectName).GetROProperty("value")
							 Call OutputCommondata(ObjectName,ObjectValue)
						 Case "LST"
							 ObjectValue = PrntObj.WebList(ObjectName).GetROProperty("value")
							 Call OutputCommondata(ObjectName,ObjectValue)	
						 Case "RDO"
							 ObjectValue = PrntObj.WebRadioGroup(ObjectName).GetROProperty("value")
							 Call OutputCommondata(ObjectName,ObjectValue)
						 Case "CHK"
							 ObjectValue = PrntObj.WebCheckbox(ObjectName).GetROProperty("value")
							 Call OutputCommondata(ObjectName,ObjectValue)
						Case "ELE"
							 ObjectValue = PrntObj.WebElement(ObjectName).GetROProperty("outertext")
							 If Instr(ObjectName,"ELE_PolicyName") Then
							 	ObjectValue = Replace(ObjectValue,"Policy Name: ","")
							 ElseIf Instr(ObjectName,"ELE_PolicyID") Then
							    ObjectValue = Replace(ObjectValue,"Policy#: ","")
								ObjectValue=Right(ObjectValue,11)
							End If
							 Call OutputCommondata(ObjectName,ObjectValue)	 
						 Case "MFO"
						       
						        gobjExcelDataAccess.Connect()
						        Dim strQuery, objTestData
						        Set objTestData = CreateObject("ADODB.Recordset")
'						        
								strQuery = "Select Field_ID from [Output_Objects$] where Obj_Name = '" & ObjectName & "'"
						        Set objTestData = gobjExcelDataAccess.ExecuteQuery(strQuery)
						        gobjExcelDataAccess.Disconnect()
						        
						        If objTestData.RecordCount = 0 Then
						            Err.Raise 3004, "DataTable Library", "Field ID is Not avialable for the field : " & ObjectName
						        End If
						        
						        Dim strDataValue, strFirstChar
						        strDataValue = Trim(objTestData(0).Value)
						
						        objTestData.Close
						        Set objTestData = Nothing
						        
						
						        If IsNull(strDataValue) Then
						            strDataValue = ""
						        End If
						 strID = strDataValue
						 Set PrntObj = TeWindow("emulator status:=Ready").TeScreen("micclass:=TeScreen").TeField("field id:="&strID)
						 ObjectValue = PrntObj.GetROProperty("text")
						 Call OutputCommondata(ObjectName,ObjectValue)
					End Select
				End If
	'Next
      
End Function

Function Outputdata(ObjName)
	Obj_Browser = gobjDataTable.GetData(Environment("strCurrentKeyword"), "Browser")
     Obj_Page = Environment("strCurrentKeyword")
'     Obj_Browser = "ICONSTDQA"
'     Obj_Page = "ICONSTDQA"
     Set PrntObj = Browser(Obj_Browser).Page(Obj_Page)
      Dim incwait
       incwait = 0
 		strType = Left(ObjName,3)
	Select Case strType
            Case "EDI"
            	Outputdata = PrntObj.WebEdit(ObjName).GetROProperty("value")
'            	OutputCommondata ObjName,Obj_innertext(1)                   
        Case "LNK"
               
         
               	  Outputdata = PrntObj.Link(ObjName).GetROProperty("innertext")
'            	OutputCommondata Replace(arrVerify(inc),"%",""),strinrtext         
            
                
        Case "ELE"
            
               	  Outputdata = PrntObj.WebElement(ObjName).GetROProperty("innertext")
'            	OutputCommondata Replace(arrVerify(inc),"%",""),strinrtext
       
        Case "LST"
        	
            	Outputdata = PrntObj.WebList(ObjName).GetROProperty("value")   	
'            	OutputCommondata Replace(Obj_innertext(0),"%",""),Obj_innertext(1) 
		
		Case Else
			Outputdata = ObjName
				
    End Select 
End Function

Function ContinueGI()
If Browser("ICON_NBV").Page("GeneralInformation").WebElement("Warning").Exist(40) Then
   Browser("ICON_NBV").Page("GeneralInformation").WebButton("BTN_Continue").Click
   End If
End Function


Function Continue()
wait(5)

End Function

Function PremiumCompare()
	strData = gobjDataTable.GetData("PremiumSumCompare", "Premium")
	strPremium =  Browser("ICON_NBV").Page("PremiumSummary").WebElement("ELE_PSP").GetROProperty("innertext")
	strAmount = Replace(strPremium,",","")
	If strData = strAmount Then
		MsgBox "Correct Premium "
		Else 
		MsgBox "Incorrect Premium"
	End If
End Function



Function Add_Eligibility_Ques()
'Wait(9)
        If Browser("ICON_NBV").Page("ClassCodeschedule").WebElement("ELE_Additional Eligibility").Exist(10) Then
		    Set ODesc = Description.Create
			ODesc("micclass").value = "WebRadioGroup"
			Set arrobj = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODesc)
			arrobjCnt = arrobj.Count
			If arrobjCnt>0 Then
				For inc = 0 To arrobjCnt-1
				arrobj(inc).Select "NO"
				Next
			End If
		
			Set ODesc= Description.Create()
			ODesc("micclass").value="WebCheckBox"
			'ODesc("innertext").value="None of the Above"
			ODesc("outerhtml").value=".*None of the Above.*"
			ODesc("checked").value = 0
			Set arrobj = Browser("ICON_NBV").Page(Environment("strCurrentKeyword")).ChildObjects(ODesc)
			'MsgBox arrobj.count
			If arrobj.Count>0 Then
				For inc = 0 To arrobj.Count-1
				arrobj(inc).Click
				Next
			End If
			Browser("ICON_NBV").Page("ClassCodeschedule").WebButton("BTN_Continue").Click
	   End If
End Function 
	 
Function CheckTime()	

 Environment("CurrentTime_1") = time()  
 If Browser("ICON_NBV").Page("NBV").WebElement("ELE_PremiumSummary").Exist(20) then
     Environment("CurrentTime_2") = time()  
  Minutes= Datediff("n",CurrentTime_1,CurrentTime_2)
 
  End if
 
 End Function
 
  
'###################################################################################################################
'Function Description   : Function to Issue a Bindable Policy
'Input Parameters		: None
'Author					: Sumith
'Date Created			: 06/25/2016
'###################################################################################################################
Sub IssuePolicy()

'wait 3
	If Browser("ICON_NBV").Page("NBV").WebButton("value:=RESERVE").Exist(10) Then
		Browser("ICON_NBV").Page("NBV").WebButton("value:=RESERVE").Click
		wait(3)
		Browser("Browser_iConnect").Page("ICON").WebButton("Continue").Click
		wait(6)
		If Browser("ICON_NBV").Page("NBV").WebButton("name:=START ISSUE").Exist(10) Then
			Browser("ICON_NBV").Page("NBV").WebButton("name:=START ISSUE").Click
			wait(3)
			Browser("Browser_iConnect").Page("ICON").WebButton("Continue").Click
		wait(9)
		End If
		Browser("ICON_NBV").Page("NBV").Link("innertext:=Edit").Click
		'Browser("ICON_NBV").Page("NBV").WebEdit("name:=AddrIDPhone").Set "(347) 821-0319"
		Browser("ICON_NBV").Page("NBV").WebEdit("name:=AddrIDPhone").Set gobjDataTable.GetData("IssuePolicy","EDI_AddrIDPhone")
		Browser("ICON_NBV").Page("NBV").WebButton("name:=Save").Click
		Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
		'Browser("ICON_NBV").Page("NBV").WebRadioGroup("name:=downPayment").Select "0"
		Browser("ICON_NBV").Page("NBV").WebRadioGroup("name:=downPayment").Select gobjDataTable.GetData("IssuePolicy","EDI_DownPayment")
		If Browser("ICON_NBV").Page("NBV").WebList("name:=BillFrequency").GetROProperty("value") <> "Full Pay" Then
			wait(2)
			Browser("ICON_NBV").Page("NBV").WebList("name:=BillFrequency").Select gobjDataTable.GetData("IssuePolicy","LST_BillFreq")
		End If
		wait(2)
		Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
		Browser("ICON_NBV").Page("NBV").Sync
		Browser("ICON_NBV").Page("NBV").WebEdit("name:=InsuredContactName").Set gobjDataTable.GetData("IssuePolicy","EDI_InsuredContactName")
		Browser("ICON_NBV").Page("NBV").WebEdit("name:=InsuredContactWorkPhone").Set gobjDataTable.GetData("IssuePolicy","EDI_InsuredContactWorkPhone")
		Browser("ICON_NBV").Page("NBV").WebEdit("name:=InsuredContactEmail").Set gobjDataTable.GetData("IssuePolicy","EDI_InsuredContactEmail")
		'Browser("ICON_NBV").Page("NBV").WebEdit("name:=NPNTextBox").Set "CHRISTMAN, BRENDA, S : 4507510"
		Browser("ICON_NBV").Page("NBV").WebEdit("name:=NPNTextBox").Set gobjDataTable.GetData("IssuePolicy","EDI_NPNTextBox")
		Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
		wait(2)
		Browser("ICON_NBV").Page("NBV").WebButton("name:=Ok").Click
		gobjReport.UpdateTestLog "Issue Policy","The Policy was issued successfully", "Pass"
'		gobjReport.UpdateTestLog "Quote ID","The Quote ID is" & QuoteID, "Fail"
	Else
		gobjReport.UpdateTestLog "Issue Policy","The Policy was not issued successfully", "Fail"
	End If
End Sub

'###################################################################################################################
'Function Description   : Function to fill in the ExpMod
'Input Parameters		: None
'Author					: Sumith
'Date Created			: 06/26/2016
'###################################################################################################################
Sub ExpMod()
	wait(5)
	If Browser("Browser_iConnect").Page("ICON").WebElement("Warning").Exist(3) Then
		Browser("ICON_NBV").Page("NBV").WebButton("BTN_Continue").Click
		wait(10)
	End If
	If Browser("ICON_NBV").Page("NBV").WebEdit("name:=experMod_0").Exist(3) Then
		Browser("ICON_NBV").Page("NBV").WebEdit("name:=experMod_0").Set "6.5"
		gobjReport.UpdateTestLog "Experience Mod","The Experience Mod was input successfully", "Pass"
'		gobjReport.UpdateTestLog "Quote ID","The Quote ID is" & QuoteID, "Fail"
	Else
		gobjReport.UpdateTestLog "Experience Mod","The Experience Mod was not inputed successfully", "Fail"
	End If
End Sub

'###################################################################################################################
'Function Description   : Function to Refer a Non-Bindable Quote to UW
'Input Parameters		: None
'Author					: Sumith
'Date Created			: 06/28/2016
'###################################################################################################################
Sub ReferUW()
	If Browser("ICON_NBV").Page("NBV").WebButton("BTN_REFER").Exist(5) Then
		Browser("ICON_NBV").Page("NBV").WebButton("BTN_REFER").Click
		wait(10)
		If Browser("ICON_NBV").Page("NBV").WebButton("BTN_referButton").Exist(3) Then
			Browser("ICON_NBV").Page("NBV").WebButton("BTN_referButton").Click
		End If
		gobjReport.UpdateTestLog "Refer UW","The Non-Bindable Quote was referred to UnderWriter successfully", "Pass"
'		gobjReport.UpdateTestLog "Quote ID","The Quote ID is" & QuoteID, "Fail"
	Else
		gobjReport.UpdateTestLog "Refer UW","The Non-Bindable Quote was not referred to UnderWriter successfully", "Fail"
	End If
End Sub

'###################################################################################################################
'Function Description   : Function to Refer a Non-Bindable Quote to UW
'Input Parameters		: None
'Author					: Sumith
'Date Created			: 06/28/2016
'###################################################################################################################
Sub EmpLiabs_Temp()
	Browser("ICON_NBV").Page("GeneralInformation").WebList("LST_EmpLiabLimits").Select "$1,000,000 / $1,000,000 / $1,000,000"
	Browser("ICON_NBV").Page("GeneralInformation").WebButton("BTN_Continue").Click
	ContinueGI
End Sub

'###################################################################################################################
'Function Description   : Function to Refer a Non-Bindable Quote to UW
'Input Parameters		: None
'Author					: Sumith
'Date Created			: 06/28/2016
'###################################################################################################################
Sub EmpLiabs()
	Browser("ICON_NBV").Page("GeneralInformation").WebList("LST_EmpLiabLimits").Select "$500,000 / $500,000 / $500,000"
	Browser("ICON_NBV").Page("GeneralInformation").WebButton("BTN_Continue").Click
	ContinueGI
End Sub

Function ReferUnderwriter()
	with Browser("ICON_NBV").Page("ReferUnderwriter")
	If .WebElement("ELE_NonBindable").Exist(1) Then
	strData = .WebElement("ELE_NonBindable").GetROProperty("innertext")
	If instr(strData,"non") Then
		.WebButton("BTN_ReferToUnderwriter").Click
		Wait 5
		.WebButton("BTN_referUW").Click
    End If 
	ElseIf  Browser("ICON_NBV").Page("ReferUnderwriter").WebElement("ELE_Bindable").Exist(1) Then
	strData = .WebElement("ELE_Bindable").GetROProperty("innertext") 
  	If instr(strData,"Bindable") Then
     	.WebList("LST_ReferUW").Select "Refer to Underwriting"
      	wait 5
     	.WebCheckBox("CHK_Refer").Set "ON"
			.WebEdit("EDI_Refer").Set "test"
			.WebButton("BTN_referUW").Click
	End If
	End If
End with
End Function

Function ExpireQuote
	Browser("AgencyPortal Console").Page("ExpireReRateQuote").Frame("menu").WebList("select").Select "Edit ACORD XML"
	XMLData1= Browser("AgencyPortal Console").Page("ExpireReRateQuote").Frame("content").WebEdit("xml_data").GetROProperty("innertext")
	If instr(XMLData1,"N158") then
	ExpireDt= Replace(XMLData1,"N186"&chr(34)&">2016","N186"&chr(34)&">2015")
	Browser("AgencyPortal Console").Page("ExpireReRateQuote").Frame("content").WebEdit("xml_data").Set ExpireDt
	Browser("AgencyPortal Console").Page("ExpireReRateQuote").Frame("content").WebButton("BTN_Update").Click
	End if
End Function




Function OutputQuoteID(strFieldName,strDataValue)
	
		Dim strFilePath, strConnectionString, r_objConn,Cmd, strQuery
		strFilePath =	"C:\Users\CO22572\Desktop\NBV\I3\I3NBV.xls"
        m_intCurrentIteration=Environment("TestIteration")
        m_intCurrentSubIteration=1
        m_strCurrentTestCase=Environment("TestName")
        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
		
		'Dim strNonQuery, objTestData
		
		strTestDataSheet = "Extract_Quote"
		
		strQuery = "Update [" & strTestDataSheet & "$] Set " & strFieldName & " = '" & strDataValue & "'" &_
																" where TC_ID = '" & m_strCurrentTestCase &_
																"' and Iteration = " & m_intCurrentIteration & ""
		
		r_objConn.Execute strQuery                       
        
        r_objConn.Close														
		
	gobjReport.UpdateTestLog "Output value", _
								"Output value " & strDataValue & " written into the " & strFieldName & " column", "Done"		
End Function

Function SearchPolicy
Dim Wshell
set Wshell = CreateObject("WScript.shell")
	strData = gobjDataTable.GetData("Extract_Quote", "Edi_PolNumber")
	Browser("ICONSmallCommercial").Page("NewCustomer").WebEdit("EDI_Search").Set strData
	Browser("ICONSmallCommercial").Page("NewCustomer").WebEdit("EDI_Search").Click
'	Wshell.SendKeys "{TAB}" 
Wshell.SendKeys "{ENTER}"
    Browser("ICONSmallCommercial").Page("NewCustomer").Link("LNK_Search").Click

End Function


Function ECOSclosebuttonhandle

if Browser("ECOSBrowser").Page("Basic Information").Link("Close").Exist then 
Browser("ECOSBrowser").Page("Basic Information").Link("Close").Click
Browser("ECOSBrowser").Page("Basic Information").WebElement("ELE_Next").Click
End If
End Function

Function ECOSpoliceditOk

'if Dialog("Message from webpage").WinButton("OK").Exist then 
Dialog("Message from webpage").WinButton("OK").Click
'End If
End Function




Function SearchCustomer
Dim Wshell
set Wshell = CreateObject("WScript.shell")
    strFieldName = "EDI_LegalName"
	strData = Getcommonoutdata(strFieldName)
	Browser("ICONSmallCommercial").Page("NewCustomer").WebEdit("EDI_CustomerName").Set strData
	Browser("ICONSmallCommercial").Page("NewCustomer").WebEdit("EDI_CustomerName").Click
'	Wshell.SendKeys "{TAB}" 
    Wshell.SendKeys "{ENTER}"
    Browser("ICONSmallCommercial").Page("NewCustomer").Link("LNK_Search").Click

End Function

Function Getcommonoutdata(strFieldName)

	    strFilePath =    "C:\NBV\e2e.xls"
	      m_intCurrentIteration=Environment("TestIteration")
	        m_intCurrentSubIteration=1
	        m_strCurrentTestCase=gobjTestParameters.CurrentTestcase
        strTestDataSheet = "EndToEnd"
        Dim strQuery, objTestData    
        
        
        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                                    ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""
        
        Set m_objConn = CreateObject("ADODB.Connection")
        m_objConn.Open strConnectionString
        Set objTestData = CreateObject("ADODB.Recordset")
        
        strQuery = "Select E2E_Fields from [Field_Mapping$] where NBV = '" &strFieldName& "'"
        
        strFieldName = ExecuteQuery(strQuery)                            
        
        strQuery = "Select " & strFieldName & " from [" & strTestDataSheet & "$]" &_
                                                " where TC_ID = '" & m_strCurrentTestCase  &_
                                                "' and Iteration = " & m_intCurrentIteration & "" 
  
         objTestData = ExecuteQuery(strQuery)
    
        Getcommonoutdata =  objTestData

End Function



Function UI_E2Evalidation(strData,strFieldName)
   UIvalue = Browser("ICONSmallCommercial").Page(Environment("strCurrentKeyword")).WebElement(strFieldName).GetROProperty("innertext") 
  If strData = "!E2E" Then
  	 E2EData = Getcommonoutdata(strFieldName)
  ElseIf strData = "!NotNull" Then
        If UIvalue <> "" Then
       gobjReport.UpdateTestLog "Comparision of NBV and PC Fields", "UI value" &UIvalue& "is not null in UI" ,"Pass" 
        Else 
       gobjReport.UpdateTestLog "Comparision of UI field value and Data provided in E2E sheet", "UI value" & UIvalue &"E2E Data value" & E2EData& "are different","Fail" 
       End If
       Exit Function
   Else
     E2EData = Replace(strData,"!","")
  End If 
  
  	   If UIvalue = E2EData Then
       gobjReport.UpdateTestLog "Comparision of NBV and PC Fields", "UI value" &UIvalue& "E2E Data value" & E2EData&"are same","Pass" 
        Else 
       gobjReport.UpdateTestLog "Comparision of NBV and PC Fields", "UI value" & UIvalue &"E2E Data value" & E2EData& "are different","Fail" 
       End If
    
End Function




Function VerificationFields(arrVerify)

	Obj_Browser = gobjDataTable.GetData(Environment("strCurrentKeyword"), "Browser")
     Obj_Page = Environment("strCurrentKeyword")

		If   Obj_Browser <> ""  Then
			Set PrntObj = Browser(Obj_Browser).Page(Obj_Page)
		End If     
      Dim incwait
        incwait = 0 	
	For inc = 1 to Ubound(arrVerify)	
	
	strVerifyField = Split(arrVerify(inc),":")
	strFieldName = Replace(strVerifyField(0),"%","")
	strFieldType = Mid(strFieldName,1,3)
	strFieldValue = strVerifyField(1)
	strFieldValueType = Mid(strFieldValue,1,3)
	
	If strFieldValueType = "EDI" or  strFieldValueType = "LST" or strFieldValueType ="CHK" or strFieldValueType ="ELE" or strFieldValueType ="RDO" Then
		   'fieldValue = Call FetchValue(strFieldValue)
		   fieldValue = Getcommonoutdata(strFieldValue)
		ElseIf Instr(1,strFieldValue,"*") Then
		   fieldValue = Replace(strFieldValue,"*","")
		Else
		   fieldValue = strFieldValue
	End If
	
	Select Case strFieldType
            Case "EDI"
              appValue = PrntObj.WebEdit(strFieldName).GetROProperty("value")
            If Instr(1,appValue,fieldValue)Then
            	PrntObj.WebEdit(strFieldName).highlight
            	gobjReport.UpdateTestLog "The Element is available in the screen"&strFieldName, "Element is sucessfuly available in screen", "Pass"
            	Else
            	gobjReport.UpdateTestLog "The Element is not available in the screen"&strFieldName, "Element is not available in screen", "Fail"
            End if	
            
           Case "LST"
              appValue = PrntObj.WebList(strFieldName).GetROProperty("value")
            If Instr(1,appValue,fieldValue)Then
            	PrntObj.WebList(strFieldName).highlight
            	gobjReport.UpdateTestLog "The Element is available in the screen"&strFieldName, "Element is sucessfuly available in screen", "Pass"
            	Else
            	gobjReport.UpdateTestLog "The Element is not available in the screen"&strFieldName, "Element is not available in screen", "Fail"
            End if		            
           
            
            
        Case "LNK"
             appValue = PrntObj.Link(strFieldName).GetROProperty("value")
            If Instr(1,appValue,fieldValue)Then
            	PrntObj.Link(strFieldName).highlight
                gobjReport.UpdateTestLog "The Element is available in the screen"&strFieldName, "Element is sucessfuly available in screen", "Pass"
            	Else
            	gobjReport.UpdateTestLog "The Element is not available in the screen"&strFieldName, "Element is not available in screen", "Fail"
            End if	
            
        Case "ELE"
             appValue = PrntObj.WebElement(strFieldName).GetROProperty("value")
            If Instr(1,appValue,fieldValue)Then
            	PrntObj.WebElement(strFieldName).highlight
             	gobjReport.UpdateTestLog "The Element is available in the screen"&strFieldName, "Element is sucessfuly available in screen", "Pass"
            	Else
            	gobjReport.UpdateTestLog "The Element is not available in the screen"&strFieldName, "Element is not available in screen", "Fail"
            End if	
         
    End Select  
    Next
	
End Function



Function FetchValue(strFieldName)
	Dim strFilePath, strConnectionString, r_objConn,Cmd, strQuery
		strFilePath =	"C:\NBV\e2e.xls"
	        m_intCurrentIteration=Environment("TestIteration")
	        m_intCurrentSubIteration=1
	        m_strCurrentTestCase=gobjTestParameters.CurrentTestcase
        

        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
        Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
		
		Dim strNonQuery, objTestData
		strTestDataSheet = "EndToEnd"
		
        strQuery = "Select E2E_Fields from [Field_Mapping$] where NBV = '" &strFieldName& "'"
        
        strFieldName = ExecuteQuery(strQuery)
        
        strQuery = "Select [" & strFieldName & "] from [EndToEnd$] where TC_ID = '" & m_strCurrentTestCase &_
                                                                "' and Iteration = " & m_intCurrentIteration & "" 
        

		
			r_objConn.Execute strQuery
	 FetchValue = ExecuteQuery(strQuery)  
		
        r_objConn.Close
        
        End Function

	


	Function GetCount()
	
		 
		strFilePath =  "C:\Forms_Xpressv2.0\Test_Data\Test_Data.xls"
        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
      
         m_strCurrentTestCase=gobjTestParameters.CurrentTestcase 
		strTestDataSheet = "FormDetails"
		strQuery = "SELECT COUNT(PolicyNumber) FROM [FormDetails$]"
		'strQuery = "Select PolicyNumber from [FormDetails$] where Field_ID = '" &FldID& "'"
		count=ExecuteQuery(strQuery) 
		GetCount=count
		'Msgbox count
		
	End Function
	
	Function RetrieveForms() 
        
        Browser("name:=.*").Page("name:=.*").WebElement("innertext:= Forms","class:=x-grid-cell-inner x-grid-cell-inner-treecolumn").Click
        strFilePath = Environment("OutputFileName")
        strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFilePath &_
                                                            ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=2"""    
      
          m_strCurrentTestCase=gobjTestParameters.CurrentTestcase 
		
		strTestDataSheet = "FormDetails"
		'm_strCurrentPolicyNumber="AA4CJ3"
		m_strexecflag="Yes"
		strfieldform= "FormName"
		strfielddesc="FormDesc"
		strPolnum="PolicyNumber"

		Set r_objConn = CreateObject("ADODB.Connection")
        r_objConn.Open strConnectionString
'         
		'Take the count fot Total no.of pages
			If Browser("name:=.*").Page("name:=.*").WebElement("innertext:=of.*","index:=0").Exist(5) Then
			  Browser("name:=.*").Page("name:=.*").WebElement("innertext:=of.*","index:=0").RefreshObject
			  wait 2 'Wait required to load the objects	
			  a=Browser("name:=.*").Page("name:=.*").WebElement("innertext:=of.*","index:=0").GetROProperty("innertext")
			  cnt=Split(a," ")
			  TotalPages=cnt(1)
			  Dim Page
			  Page=1	
			Else
			  TotalPages = 1		
			End If
		
			'FormsVal=Split(strValue,"::")
			If TotalPages=1 Then
					Set objParent=Browser("name:=.*").Page("title:=.*")
					Set objTable = objParent.WebTable("html id:=gridview.*","class:=x-gridview.*","index:=0")
					objTable.RefreshObject
					Rowcnt=objTable.RowCount
						For i  = 1 To Rowcnt
							FormN=objTable.GetCellData(i,1)
							FormD=objTable.GetCellData(i,2)
							Environment("FieldID")=GetCount()
							 FldID = "Field"&Environment("FieldID")
							'Verify whether the table value is equal to input data
							strQuery = "Update [" & strTestDataSheet & "$] Set " & strfieldform & " = '" & FormN & "'" &_
																" where  Field_ID = '"& FldID &"'" 
						r_objConn.Execute strQuery 
						
		
       					 r_objConn.Close
       					' strQuery = "Update [" & strTestDataSheet & "$] Set " & strfielddesc & " = '" & FormD & "'" &_
																'" where  Field_ID = '"& FldID &"'" 
                                                              '  r_objConn.Execute strQuery 
						
		
       					 r_objConn.Close
						Next
			Else
				For Iterator = 1 To TotalPages
					Set objParent=Browser("name:=.*").Page("title:=.*")
					Set objTable = objParent.WebTable("html id:=gridview.*","class:=x-gridview.*","index:=0")
					Rowcnt=objTable.RowCount
						For i  = 1 To Rowcnt
							FormN=objTable.GetCellData(i,1)
							FormD=objTable.GetCellData(i,2)
							Environment("FieldID")=GetCount()
							'Verify whether the table value is equal to input data
							FldID = "Field"&Environment("FieldID")
							
							strQuery = "Update [" & strTestDataSheet & "$] Set " & strfieldform & " = '" & FormN & "'," & strPolnum  & "='" & Environment("Policy") &"'" &_
																			" where  Field_ID = '"& FldID  &"'" 
						
		
       					' r_objConn.Close
       					 'strQuery = "Update [" & strTestDataSheet & "$] Set " & strfielddesc & " = '" & FormD & "'" &_
																'" where  Field_ID = '"& FldID &"'" 
                         r_objConn.Execute strQuery 
						
		'Environment("FieldID") = Environment("FieldID") + 1
       					' r_objConn.Close
						Next
						Page=cint(Page)+1
						'Click on next arrow
						objParent.WebElement("html id:=gbutton-.*","innerhtml:=&nbsp;","Class:=x-btn-icon-el x-tbar-page-next ").Click
						wait 2 'Wait required to load the objects	
						
				Next
				
		End If
	    		
	End Function        

	
	
Function GenerateRandomString(StrLength, strtype)

 ' Declare Variables
Dim StrLenIndex
Dim chrAscCode
Dim gChr

If strtype = "Text" Then

	 'Use for loop to generate StrLength many characters
	For StrLenIndex = 1 To Round(StrLength/3)
	
	  ' Get Random ASCII Code
	  chrAscCode=RandomNumber(65,90)
	  ' Convert ASCII code in to character
	  gChr=gChr&Chr(chrAscCode)
	Next
	gChr=gChr& " "
	For StrLenIndex = 1 To Abs(StrLength/3)
	  ' Get Random ASCII Code
	  chrAscCode=RandomNumber(48,57)
	  ' Convert ASCII code in to character
	  gChr=gChr & Chr(chrAscCode)
	  
	Next  
	 gChr=gChr& " "
	For StrLenIndex = 1 To Abs(StrLength/3)
	
	  ' Get Random ASCII Code
	  chrAscCode=RandomNumber(65,90)
	  ' Convert ASCII code in to character
	  gChr=gChr&Chr(chrAscCode)
	  
	Next
	  
	'Return converted character to function
	
	
	Else
	
	For StrLenIndex = 1 To StrLength
	  ' Get Random ASCII Code
	  chrAscCode=RandomNumber(48,57)
	  ' Convert ASCII code in to character
	  gChr=gChr & Chr(chrAscCode)
	  
	Next  
	 
	
End If

GenerateRandomString=gChr

End Function

Function RTS
	


Const forReading = 1
Const forWriting = 2
Const forAppend = 8 

Datatable.AddSheet "Sheet1"
Datatable.ImportSheet "\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2017\0311 Release\Test Execution\CDS\Regression\functionaltesting.xls",1,"Sheet1"

For i = 1 To datatable.GetSheet("Sheet1").GetRowCount
datatable.GetSheet("Sheet1").SetCurrentRow(i)
Policynumber = datatable.Value ("Policynumber","Sheet1")
print Policynumber
PPM =	datatable.Value ("PPM","Sheet1")
Print PPM
path =datatable.Value("Path","Sheet1")
Print path
newvalue2 =  datatable.Value ("Response","Sheet1")
print newvalue2
newvalue =  datatable.Value ("Request","Sheet1")
print newvalue
datatable.SetNextRow
strfolderpath = path & "\" &  PPM
 
Set obj =CreateObject("Scripting.FileSystemObject")
If obj.FolderExists(strfolderpath) =false Then
obj.CreateFolder strfolderpath
End If
Set obj = nothing 

SystemUtil.Run "iexplore", "http://qa01ebc.thehartford.com"
Browser("Login | The Hartford EBC").Page("Login | The Hartford EBC").WebEdit("USERText").Set "nbvqa197"
Browser("Login | The Hartford EBC").Page("Login | The Hartford EBC").WebEdit("PASSWORD").SetSecure "58cf5bc44991da4f7fa6d80a4ac74ed3616b7f6ade31"
Browser("Login | The Hartford EBC").Page("Login | The Hartford EBC").Image("Login").Click 2,2
Browser("Login | The Hartford EBC").Page("Commercial Insurance:").Link("Quote Small Commercial").Click
wait (40)

Browser("ICON | Small Commercial").Page("ICON | Small Commercial").WebList("selectSearch").Select "Policy #"
wait(5)
Set oshell = createobject("WScript.Shell")


Browser("ICON | Small Commercial").Page("ICON | Small Commercial").WebEdit("keywordPolicyNumber").Click
oshell.SendKeys(Policynumber)
Browser("ICON | Small Commercial").Page("ICON | Small Commercial").Link("Search").Click
wait(5)
Browser("iConnect").Page("ICON | Small Commercial").WebTable("Customer").ChildItem(2,4,"Link",0).click

wait(5)
Browser("AgencyPortal Console").Page("AgencyPortal Console").Frame("AgencyPortal Console Menu").WebList("select").Select "RTS Standard Request"
wait(5)
a= Browser("AgencyPortal Console").Page("AgencyPortal Console").Frame("Frame").WebEdit("xml_data").GetROProperty("innertext")

Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName = "\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml"
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Write a
oFile.Close
Browser("Login | The Hartford EBC").Page("Commercial Insurance:").Link("Log Off").Click
systemutil.CloseProcessByName "iexplore.exe"

Set doc = XMLUtil.CreateXML()
newdata = "\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml"
doc.LoadFile(newdata)
doc.SaveFile "\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml"
Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName = strfolderpath & "\" &  newvalue
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Close

Set fso = CreateObject("Scripting.FileSystemObject")
Set strnewtext = fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml" , 1, false)
strnewtext.SkipLine
x=strnewtext.ReadAll
strnewtext.Close
Set strnewtext = fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml", forWriting)
strnewtext.Write x
strnewtext.Close

Set strnewtext= fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml", 1, false)
strnewtext.SkipLine
x=strnewtext.ReadAll
strnewtext.Close
Set strnewtext = fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml", forWriting)
strnewtext.Write x
strnewtext.Close


Set fso = CreateObject("Scripting.FileSystemObject")
Set objfile3 = fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2017\0311 Release\Test Execution\CDS\Regression\SoapPart1.xml", forReading)
strtext=objfile3.ReadAll
Set objfile2 = fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\0910 Release\Test Execution\CDS\Cycle 3\newl2.xml", forReading)
strtextnew=objfile2.ReadAll


Set objfile1 = fso.openTextfile("\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2017\0311 Release\Test Execution\CDS\Regression\SoapPart2.xml", forReading)
strtext1=objfile1.ReadAll
Set objfiles = fso.OpenTextFile(strfolderpath & "\" &  newvalue, forAppend)
objfiles.WriteLine strtext
objfiles.WriteLine strtextnew 
objfiles.WriteLine strtext1
objfiles.Close
objfile1.Close
objfile2.Close
objfile3.Close



Dim XMLDataFile2 
XMLDataFile2 = strfolderpath & "\" &  newvalue
Set xmlDoc2 = CreateObject("Microsoft.XMLDOM")
xmlDoc2.Async= False
b=xmlDoc2.Load(XMLDataFile2)
Set currentnode =xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/htfd:LastAccessDt ") 
For each Accessdate in currentnode 
Accessdate.ParentNode.RemoveChild(Accessdate)
Next
Set accountmangement = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AccountManagementProcessInd")
For each accountmangamentprocess in accountmangement 
accountmangamentprocess.ParentNode.RemoveChild(accountmangamentprocess)
Next
Set accountmangement1 = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AgentDifferentiatedBindabilityFlagApprovedInd")
For each objnode7 in accountmangement1 
objnode7.ParentNode.RemoveChild(objnode7)
Next
Set accountmangement2 = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AgentDifferentiatedBindabilityQuoteInd")
For each objnode8 in accountmangement2 
objnode8.ParentNode.RemoveChild(objnode8)
Next
Set accountmangement23 = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:UpsellInfo")
For each objnode81 in accountmangement23 
objnode81.ParentNode.RemoveChild(objnode81)
Next
Set UpsellReportingInfo = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:UpsellReportingInfo")
For each objnod in UpsellReportingInfo 
objnod.ParentNode.RemoveChild(objnod)
Next
Set currentnode =xmlDoc2.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:ClassData") 
Set soapnew = xmlDoc2.createElement("htfd:BridgeRequestInd")
Set soaptext = xmlDoc2.createTextNode(1)
soapnew.AppendChild soaptext
currentnode.parentNode.InsertBefore  soapnew,currentnode  
xmlDoc2.Save(XMLDataFile2) 

Set objfiln = fso.openTextfile(strfolderpath & "\" &  newvalue, forReading)
strLine=objfiln.ReadAll

SystemUtil.Run "iexplore", "http://wssemcicontentsint.thehartford.com/RTSemciContents/SOA-Tester/RTS-Tester.html"
Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebList("environments").Select "QA - RTS"
Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebEdit("xmlRequest").Set(strLine)
Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebButton("Submit RTS Soap Request").Click
Dialog("Internet Explorer").WinButton("Yes").Click
wait(50)
b = Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebEdit("xmlResponse").GetROProperty("innertext")
Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName = strfolderpath & "\" &  newvalue2
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Write b
ofile.Close
 
systemutil.CloseProcessByName "iexplore.exe"




Dim XMLDataFile3 
XMLDataFile3 = strfolderpath & "\" &  newvalue
Set xmlDoc3 = CreateObject("Microsoft.XMLDOM")
xmlDoc3.Async= False
c=xmlDoc3.Load(XMLDataFile3)
Set currentnode =xmlDoc3.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo")
Set other = xmlDoc3.createElement("OtherIdentifier")
currentnode.AppendChild(other) 

Set currentnode4 =xmlDoc3.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier")
Set other1 = xmlDoc3.createElement("OtherIdTypeCd")
other1.Text ="com.thehartford_NewIconCustId"
currentnode4.AppendChild(other1) 

Set currentnode5 =xmlDoc3.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier")
Set other2 = xmlDoc3.createElement("OtherId")
other2.Text ="0008114579"
currentnode5.AppendChild(other2) 

Set currentnode7 =xmlDoc3.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo")
Set other7 = xmlDoc3.createElement("OtherIdentifier")
currentnode7.AppendChild(other7)

Set currentnode8 =xmlDoc3.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier[1]")
Set other8 = xmlDoc3.createElement("OtherIdTypeCd")
other8.Text ="com.thehartford_CustId"
currentnode8.AppendChild(other8) 

Set currentnode10 =xmlDoc3.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier[1]")
Set other10 = xmlDoc3.createElement("OtherId")
other10.Text ="000811"
currentnode10.AppendChild(other10) 

xmlDoc3.Save(XMLDataFile3) 

Set fso = CreateObject("Scripting.FileSystemObject")
Set objfile = fso.openTextfile(strfolderpath & "\" &  newvalue, forReading)

strtext=objfile.ReadAll
objfile.Close
strnewtext = Replace(strtext, "OtherIdentifier xmlns", "OtherIdentifier text")
Set objfile = fso.openTextfile(strfolderpath & "\" &  newvalue, forWriting)
objfile.WriteLine strnewtext
objfile.Close

Dim XMLDataFile12 
XMLDataFile12 = strfolderpath & "\" &  newvalue
Set xmlDoc12= CreateObject("Microsoft.XMLDOM")
xmlDoc12.Async= False
c=xmlDoc12.Load(XMLDataFile12)


Set otheridentifier = xmlDoc12.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier")
For each  otheridentifiers in  otheridentifier
otheridentifiers.RemoveAttribute("text")
Next
xmlDoc12.save(XMLDataFile12)


Dim XMLDataFileRequote 
XMLDataFileRequote  = strfolderpath & "\" &  newvalue
Set xmlDocRequote  = CreateObject("Microsoft.XMLDOM")
xmlDocRequote.Async= False
c=xmlDocRequote.Load(XMLDataFileRequote)
Set currentnodeRequote =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AutoOnPolCd")

Set soapnewRequote = xmlDocRequote.createElement("QuoteInfo")
currentnodeRequote.parentNode.InsertBefore  soapnewRequote,currentnodeRequote

Set currentnodeRequote245 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
Set otherRequote245 = xmlDocRequote.createElement("CompanysQuoteNumber")
otherRequote245.Text ="2017-03-11"
currentnodeRequote245.AppendChild(otherRequote245)


Set currentnode51 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
Set otherRequote2 = xmlDocRequote.createElement("ItemIdInfo")
currentnode51.AppendChild(otherRequote2) 

Set currentnodeRequote26 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo/ItemIdInfo")
Set otherRequote26 = xmlDocRequote.createElement("InsurerId")
otherRequote26.Text ="001"
currentnodeRequote26.AppendChild(otherRequote26)

Set currentnodeRequote24 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
Set otherRequote24 = xmlDocRequote.createElement("InitialQuoteRequestDt")
otherRequote24.Text ="2017-03-11"
currentnodeRequote24.AppendChild(otherRequote24)



Set currentnodeRequote28 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
Set otherRequote28 = xmlDocRequote.createElement("Description")
otherRequote28.Text ="Original"
currentnodeRequote28.AppendChild(otherRequote28)

Set currentnodeRequote30 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
Set otherRequote30 = xmlDocRequote.createElement("htfd:OriginalPolicyNumber")
otherRequote30.Text ="02SB 4996IE"
currentnodeRequote30.AppendChild(otherRequote30)

Set currentnodeRequote32 =xmlDocRequote.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
Set otherRequote32 = xmlDocRequote.createElement("htfd:OverrideInitialQuoteRequestDt")
otherRequote32.Text ="2017-03-11"
currentnodeRequote32.AppendChild(otherRequote32)

xmlDocRequote.Save(XMLDataFileRequote) 

Set fso12 = CreateObject("Scripting.FileSystemObject")
Set objfile12 = fso12.openTextfile(strfolderpath & "\" &  newvalue, forReading)

strtext12=objfile12.ReadAll

objfile12.Close
Strnewtext12 = Replace(strtext12, "QuoteInfo xmlns", "QuoteInfo text")
Set objfile12 = fso12.openTextfile(strfolderpath & "\" &  newvalue, forWriting)
objfile12.WriteLine Strnewtext12
objfile12.Close

Dim XMLDataFile13 
XMLDataFile13 = strfolderpath & "\" &  newvalue
Set xmlDoc13= CreateObject("Microsoft.XMLDOM")
xmlDoc13.Async= False
c=xmlDoc13.Load(XMLDataFile13)


Set otheridentifier13 = xmlDoc13.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
For each  otheridentifiers13 in  otheridentifier13
otheridentifiers13.RemoveAttribute("text")
Next
	
xmlDoc13.save(XMLDataFile13)

Dim XMLDataFilenew 
XMLDataFilenew = strfolderpath & "\" &  newvalue
Set xmlDoc13= CreateObject("Microsoft.XMLDOM")
xmlDoc13.Async= False
c=xmlDoc13.Load(XMLDataFile13)

Set otheridentifier13 = xmlDoc13.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
For each  otheridentifiers13 in  otheridentifier13
otheridentifiers13.RemoveAttribute("text")
Next

Set otheridentifier14 = xmlDoc13.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo")
For each  otheridentifiers14 in  otheridentifier14
otheridentifiers14.RemoveAttribute("text")
Next
 
xmlDoc13.save(XMLDataFilenew)


Set doc = XMLUtil.CreateXML()
newdata =  strfolderpath & "\" &  newvalue
doc.LoadFile(newdata)
doc.SaveFile (strfolderpath & "\" &  newvalue)


Dim XMLDataFilenew13 
XMLDataFilenew13  = strfolderpath & "\" &  newvalue2
Set xmlDocnewtest = CreateObject("Microsoft.XMLDOM")
xmlDocnewtest.Async= False
test1=xmlDocnewtest.Load(XMLDataFilenew13)
msgbox test1
 
Set nodes =xmlDocnewtest.SelectNodes("/soapenv:Envelope/soapenv:Body/ACORD/InsuranceSvcRs/BOPPolicyQuoteInqRs/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier/OtherId/text()")
For a = 0 To (nodes.Length -1)
policy = nodes(i).NodeValue
Next
msgbox "otheridentifier number :  " & policy



Set nodes16 =xmlDocnewtest.SelectSingleNode("/soapenv:Envelope/soapenv:Body/ACORD/InsuranceSvcRs/BOPPolicyQuoteInqRs/CommlPolicy/QuoteInfo/CompanysQuoteNumber/text()")

nodenew = nodes16.NodeValue 

msgbox "Policynumber number :  " & nodenew
 
 
xmlDocnewtest.save(XMLDataFilenew13)

Dim XMLDataFilenewtest 
XMLDataFilenewtest  = strfolderpath & "\" &  newvalue
Set XMLDataFilenewtestnew = CreateObject("Microsoft.XMLDOM")
XMLDataFilenewtestnew.Async= False
b=XMLDataFilenewtestnew.Load(XMLDataFilenewtest)
msgbox b

Dim nodes
Set nodes2 =XMLDataFilenewtestnew.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo/CompanysQuoteNumber") 
For  each objnode in nodes2 
objnode.Text = nodenew
Next

Set nodes3 =XMLDataFilenewtestnew.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/QuoteInfo/htfd:OriginalPolicyNumber") 
For  each objnode1 in nodes3 
objnode1.Text = nodenew
Next
   
Set nodes4 =XMLDataFilenewtestnew.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier/OtherId") 
For  each objnode2 in nodes4 
objnode2.Text = policy
Next
XMLDataFilenewtestnew.Save(XMLDataFilenewtest)  
    

Set fso1 = CreateObject("Scripting.FileSystemObject")
Set fpath1 = fso1.openTextfile(strfolderpath & "\" &  newvalue, forReading)
strLine1 = fpath1.ReadAll
SystemUtil.Run "iexplore","http://wssemcicontentsint.thehartford.com/RTSemciContents/SOA-Tester/RTS-Tester.html"


Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebList("environments").Select "QA - RTS"

Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebEdit("xmlRequest").Set(strLine1)
Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebButton("Submit RTS Soap Request").Click
Dialog("Internet Explorer").WinButton("Yes").Click
new1 = Browser("Login | The Hartford EBC").Page("RTS - Test Client").WebEdit("xmlResponse").GetROProperty("innertext")

Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName = strfolderpath & "\" &  newvalue2
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Write new1
ofile.Close
 
systemutil.CloseProcessByName "iexplore.exe"

Next

End Function	
	
Function PrepareRTSXML()

		Const forReading = 1
		Const forWriting = 2
		Const forAppend = 8 
		
		
		
'		Policynumber = gobjDataTable.GetData("RTS", "Policynumber")
'		print Policynumber
		PPM = gobjTestParameters.CurrentTestcase
		Print PPM
		path = gobjDataTable.GetData("RTS", "Path") 
		Print path
		newvalue2 =  gobjDataTable.GetData("RTS", "Response") 
		print newvalue2
		newvalue =  gobjDataTable.GetData("RTS", "Request") 
		print newvalue
		datatable.SetNextRow
		strfolderpath = path & "\" &  PPM
		 
		Set obj =CreateObject("Scripting.FileSystemObject")
		If obj.FolderExists(strfolderpath) =false Then
		obj.CreateFolder strfolderpath
		End If
		Set obj = nothing 



a= Browser("AgencyPortal Console").Page("AgencyPortalConsole").WebEdit("EDI_xml_data_rts").GetROProperty("innertext")

Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName = "H:\Muthu Scripts\RTS\XML\newl2.xml"
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Write a
oFile.Close

Set doc = XMLUtil.CreateXML()
newdata = "H:\Muthu Scripts\RTS\XML\newl2.xml"
doc.LoadFile(newdata)
doc.SaveFile "H:\Muthu Scripts\RTS\XML\newl2.xml"
Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName = strfolderpath & "\" &  newvalue & ".xml"
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Close

Set fso = CreateObject("Scripting.FileSystemObject")
Set strnewtext = fso.openTextfile("H:\Muthu Scripts\RTS\XML\newl2.xml" , 1, false)
strnewtext.SkipLine
x=strnewtext.ReadAll
strnewtext.Close
Set strnewtext = fso.openTextfile("H:\Muthu Scripts\RTS\XML\newl2.xml", forWriting)
strnewtext.Write x
strnewtext.Close

Set strnewtext= fso.openTextfile("H:\Muthu Scripts\RTS\XML\newl2.xml", 1, false)
strnewtext.SkipLine
x=strnewtext.ReadAll
strnewtext.Close
Set strnewtext = fso.openTextfile("H:\Muthu Scripts\RTS\XML\newl2.xml", forWriting)
strnewtext.Write x
strnewtext.Close


Set fso = CreateObject("Scripting.FileSystemObject")
Set objfile3 = fso.openTextfile("H:\Muthu Scripts\RTS\XML\Base\SoapPart1.xml", forReading)
strtext=objfile3.ReadAll
Set objfile2 = fso.openTextfile("H:\Muthu Scripts\RTS\XML\newl2.xml", forReading)
strtextnew=objfile2.ReadAll


Set objfile1 = fso.openTextfile("H:\Muthu Scripts\RTS\XML\Base\SoapPart2.xml", forReading)
strtext1=objfile1.ReadAll
'Set objfiles = fso.OpenTextFile(strfolderpath & "\" &  newvalue & ".xml", forAppend)
Set objfiles = fso.OpenTextFile(oFileName, forAppend)

objfiles.WriteLine strtext
objfiles.WriteLine strtextnew 
objfiles.WriteLine strtext1
objfiles.Close
objfile1.Close
objfile2.Close
objfile3.Close



Dim XMLDataFile2 
XMLDataFile2 = strfolderpath & "\" &  newvalue & ".xml"
Set xmlDoc2 = CreateObject("Microsoft.XMLDOM")
xmlDoc2.Async= False
b=xmlDoc2.Load(XMLDataFile2)
Set currentnode =xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/InsuredOrPrincipal/htfd:LastAccessDt ") 
For each Accessdate in currentnode 
Accessdate.ParentNode.RemoveChild(Accessdate)
Next
Set accountmangement = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AccountManagementProcessInd")
For each accountmangamentprocess in accountmangement 
accountmangamentprocess.ParentNode.RemoveChild(accountmangamentprocess)
Next
Set accountmangement1 = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AgentDifferentiatedBindabilityFlagApprovedInd")
For each objnode7 in accountmangement1 
objnode7.ParentNode.RemoveChild(objnode7)
Next
Set accountmangement2 = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:AgentDifferentiatedBindabilityQuoteInd")
For each objnode8 in accountmangement2 
objnode8.ParentNode.RemoveChild(objnode8)
Next
Set accountmangement23 = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:UpsellInfo")
For each objnode81 in accountmangement23 
objnode81.ParentNode.RemoveChild(objnode81)
Next
Set UpsellReportingInfo = xmlDoc2.SelectNodes("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:UpsellReportingInfo")
For each objnod in UpsellReportingInfo 
objnod.ParentNode.RemoveChild(objnod)
Next
Set currentnode =xmlDoc2.SelectSingleNode("soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/BOPPolicyQuoteInqRq/CommlPolicy/htfd:ClassData") 
Set soapnew = xmlDoc2.createElement("htfd:BridgeRequestInd")
Set soaptext = xmlDoc2.createTextNode(1)
soapnew.AppendChild soaptext
currentnode.parentNode.InsertBefore  soapnew,currentnode  
xmlDoc2.Save(XMLDataFile2) 

'////launch RTS tester and submit the request xml
Set xmldoc = createobject ("Microsoft.XMLDOM")
xml_path = oFileName
Set doc = XMLUtil.CreateXML()
Set objxml = XMLUtil.CreateXMLFromFile(xml_path)
strxml = objxml.ToString
theURL = "https://rtsqa.thehartford.com/RTSService"
Set winhttp = CreateObject("microsoft.xmlhttp")
winhttp.Open "POST", theURL, False
winhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
winhttp.send strxml
responsetext = winhttp.responseText

'////paste response xml in the below given path 
Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName ="H:\Muthu Scripts\RTS\XML\Request\RTS\Auto_WithDown_QQ Rs.xml"
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Write responsetext
ofile.Close


'////Copy quoteinfo and itemidinfo tags from response XML
Const XMLDataFile = "H:\Muthu Scripts\RTS\XML\Request\RTS\Auto_WithDown_QQ Rs.xml"
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async= False
b=xmlDoc.Load(XMLDataFile)
Set nodes =xmlDoc.SelectNodes("/soapenv:Envelope/soapenv:Body/ACORD/InsuranceSvcRs/CommlAutoPolicyQuoteInqRs/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier/OtherId/text()")
For j = 0 To (nodes.Length -1)
policy = nodes(i).NodeValue
Next
Set nodes =xmlDoc.SelectNodes("/soapenv:Envelope/soapenv:Body/ACORD/InsuranceSvcRs/CommlAutoPolicyQuoteInqRs/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier[2]/OtherId/text()")
For n = 0 To (nodes.Length -1)
policycenter = nodes(i).NodeValue
Next
Set nodes1 =xmlDoc.SelectNodes("/soapenv:Envelope/soapenv:Body/ACORD/InsuranceSvcRs/CommlAutoPolicyQuoteInqRs/CommlPolicy/QuoteInfo/CompanysQuoteNumber/text()")
For i = 0 To (nodes1.Length -1)
polic1y = nodes1(i).NodeValue
Next  
xmlDoc.save(XMLDataFile)


    
  ''//// load Quick quote request XML from the given path and paste copied tags value   
 'Const XMLDataFile2 = "\\ad1.prod\HIG\BusIns\RSD\QA_Sourcing\CTS\Access\BIQA\CTS\BI QA - Offshore\BI Projects\MTC Projects\2016\1210 Release\Test Execution\CDS\Smoketest\Automation\Auto\Auto_WithDown_RiRQ.xml"
 Set xmlDoc2 = CreateObject("Microsoft.XMLDOM")
 xmlDoc2.Async= False
 b=xmlDoc2.Load(XMLDataFile2)
 Dim nodes
 Set nodes2 =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/CommlPolicy/QuoteInfo/CompanysQuoteNumber") 
 For  each objnode in nodes2 
 objnode.Text = polic1y
 Next
 Set nodes3 =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/CommlPolicy/QuoteInfo/htfd:OriginalPolicyNumber") 
 For  each objnode1 in nodes3 
 objnode1.Text = polic1y
 Next
 Set nodes4 =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier/OtherId") 
 For  each objnode2 in nodes4 
 objnode2.Text = policy
 Next 
 Set nodespc =xmlDoc2.SelectSingleNode("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/InsuredOrPrincipal/ItemIdInfo/OtherIdentifier[2]/OtherId") 
 nodespc.Text = policycenter
 Set rquidrequest =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/RqUID")
 For  each objnode2 in rquidrequest 
 objnode2.Text = rquidvalue
 Next
 Set nodes10 =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/InsuredOrPrincipal/GeneralPartyInfo/NameInfo/CommlName/CommercialName")
 Set nodes11 =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/CommlPolicy/ContractTerm/EffectiveDt") 
 Set nodes12 =xmlDoc2.SelectNodes("/soap:Envelope/soap:Body/ACORD/InsuranceSvcRq/CommlAutoPolicyQuoteInqRq/CommlPolicy/ContractTerm/ExpirationDt")  
 For  each objnode in nodes10
 objnode.Text = commercialname
 Next
 For  each objnode in nodes11
 objnode.Text = effectivedate
 Next   
 For  each objnode in nodes12
 objnode.Text = expirationdt
 Next
 xmlDoc2.Save(XMLDataFile2)  
    


'////launch RTS tester and submit the request xml

Set xmldoc = createobject ("Microsoft.XMLDOM")
xml_path_request = XMLDataFile2
Set doc = XMLUtil.CreateXML()
Set objxml = XMLUtil.CreateXMLFromFile(xml_path_request)
strxmlrequest = objxml.ToString
theURL = "https://rtsqa.thehartford.com/RTSService"
Set winhttp_request = CreateObject("microsoft.xmlhttp")
winhttp_request.Open "POST", theURL, False
winhttp_request.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
winhttp_request.send strxmlrequest
responsetext2 = winhttp_request.responseText

'////paste response xml in the below given path 
Set oFSO =CreateObject("Scripting.FileSystemObject")
oFileName ="H:\Muthu Scripts\RTS\XML\Request\RTS\Auto_WithDown_RiRs_new.xml"
Set oFile = oFSO.CreateTextFile(oFileName, True)
oFile.Write responsetext2
ofile.Close
End Function	
	
	
	
	
	
