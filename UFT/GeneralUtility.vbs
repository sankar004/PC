'General Header
'#######################################################################################################################
'Script Description		: Framework utility library
'Test Tool/Version		: HP Quick Test Professional 10+
'Test Tool Settings		: N.A.
'Application Automated	: Flight Application
'Author					: Cognizant
'Date Created			: 30/07/2008
'#######################################################################################################################
Option Explicit	'Forcing Variable declarations

Dim gobjUtil: Set gobjUtil = New Util

'#######################################################################################################################
'Class Description   	: Class to encapsulate utility functions of the framework
'Author					: Cognizant
'Date Created			: 23/07/2012
'#######################################################################################################################
Class Util
	
	'###################################################################################################################
	'Function Description   : Function to calculate the execution time for the current iteration	
	'Input Parameters 		: dtmIteration_StartTime, dtmIteration_EndTime	
	'Return Value    		: sngIteration_ExecutionTime	
	'Author					: Cognizant	
	'Date Created			: 23/01/2009	
	'###################################################################################################################
	Public Function GetTimeDifference(dtmIteration_StartTime, dtmIteration_EndTime)
		Dim strSeconds, strMinutes
		Dim sngIteration_ExecutionTime
		sngIteration_ExecutionTime = DateDiff("s", dtmIteration_StartTime, dtmIteration_EndTime)
		sngIteration_ExecutionTime = CSng(sngIteration_ExecutionTime)
		strSeconds = sngIteration_ExecutionTime Mod 60
		If Len(strSeconds) = 1 Then
			strSeconds = "0" & strSeconds
		End If
		strMinutes = (sngIteration_ExecutionTime Mod 3600) \ 60
		GetTimeDifference = strMinutes & " minute(s), " & strSeconds & " seconds"
	End Function	
	'###################################################################################################################
	
End Class
'#######################################################################################################################