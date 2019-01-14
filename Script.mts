'################################################################################### @@ hightlight id_;_65924_;_script infofile_;_ZIP::ssf50.xml_;_
'====================================================================================
'*
' Version of application
'*
'====================================================================================
SystemUtil.Run "C://AWRoot/bin/launcher/LctPOLUX.exe" 
wait 15
mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
While InStr(1, mytext, "Please wait") <> 0
	mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
Wend
wait 10
mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
While InStr(1, mytext, "Do you search for new updates?") <> 0
	mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
	If InStr(1, mytext, "Do you search for new updates?") <> 0 Then
		Browser("Home").Page("Home").WebButton("CANCEL_BTN").Click			
	End If
Wend
mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
If InStr(1, mytext, "v09.34") <> 0 Then
	Desktop.CaptureBitmap "report_image.png", True
	Reporter.ReportEvent micDone, "Information", "The version of DiagBox is v09.38", "report_image.png"
Else
	Desktop.CaptureBitmap "report_image.png", True
	Reporter.ReportEvent micFail, "Information", "The version of DiagBox is not v09.38", "report_image.png"
	SystemUtil.CloseProcessByName "diagnostic.exe"
End If
'####################################################################################









'#################################################################################### @@ hightlight id_;_10000000_;_script infofile_;_ZIP::ssf32.xml_;_
'====================================================================================
'*
' Recognition of text
'*
'==================================================================================
'Reporter.Filter = 3
'Browser("Home").Page("Home").WebButton("jeton_user").Click
'wait 2
'While InStr(1, mytext, "AUTHENTICATION") = 0
'		mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'Wend
'
'If InStr(1, mytext, "AUTHENTICATION") <> 0  Then	
'	'Browser("Home").Page("Home").WebButton("select").Click
'	'Browser("Home").Page("Home").WebButton("select").ChildObjects "DS"
'	Browser("Home").Page("Home").WebEdit("UserName").Set "E518766rr6r20" @@ hightlight id_;_10000000_;_script infofile_;_ZIP::ssf38.xml_;_
'	Browser("Home").Page("Home").WebEdit("UserPass").Set "Cst78788"
'	Browser("Home").Page("Home").WebButton("OK_BTN").Click
'	wait 2
'	mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'	While InStr(1, mytext, "Please wait") <> 0 and InStr(1, mytext, "AUTHENTICATION ERROR") = 0
'		mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'	Wend
'	mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'	If InStr(1, mytext, "AUTHENTICATION ERROR") <> 0  Then
'		Reporter.Filter = 0
'		Desktop.CaptureBitmap "report_image.png", True
'		Reporter.ReportEvent micFail, "AUTHENTICATION ERROR", "Incorrect identifier", "report_image.png"
'		Browser("Home").Page("Home").WebButton("CANCEL_BTN").Click
'	Else
'		Reporter.Filter = 0
'		Desktop.CaptureBitmap "report_image.png", True
'		Reporter.ReportEvent micDone, "Success", "Correct identifier", "report_image.png"
'	End If		
'Else
'	Reporter.Filter = 0
'	Desktop.CaptureBitmap "report_image.png", True
'	Reporter.ReportEvent micWarn, "No AUTHENTICATION pop-up", "No AUTHENTICATION pop-up", "report_image.png"
'End If
'####################################################################################		









'#################################################################################### @@ hightlight id_;_10000000_;_script infofile_;_ZIP::ssf32.xml_;_
'====================================================================================
'*
' Recognition of Authentification button
'*
'==================================================================================
'Set obj = CreateObject("Mercury.DeviceReplay")
'
'mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'comptext = "AUTHENTICATION"
'strmatch = InStr(1, mytext, comptext)
'If 	InStr(1, mytext, "Select the marque") <> 0 Then
'	If strmatch <> 0  Then
'		Desktop.CaptureBitmap "report_image.png", True
'		Reporter.ReportEvent micDone, "Success", "The text matches!" & mytext, "report_image.png"
'		Browser("Home").Page("Home").WebButton("CANCEL_BTN_2").Click
'		'obj.PressKey 14
'		Wait 10
'	Else 
'	Desktop.CaptureBitmap "report_image.png", True
'	Reporter.ReportEvent micFail, "Error", "The text doesn't matches!", "report_image.png"
'	End If
'End  If
'####################################################################################










'####################################################################################
'====================================================================================
'*
' Open the application
'*
'====================================================================================

'On Error Resume Next
'Reporter.Filter = 3
'SystemUtil.Run "C://AWRoot/bin/launcher/LctPOLUX.exe" 
'If Err.Number <> 0 Then
'	Reporter.Filter = 0
'	Desktop.CaptureBitmap "report_image.png", True
'    Reporter.ReportEvent micFail, "Failled", "The path is invalid!", "report_image.png"
'Else
'	If 	Browser("Home").Page("Home").WebButton("jeton_user").Exist(50) Then
'		Reporter.Filter = 0
'		Desktop.CaptureBitmap "report_image.png", True
'		Reporter.ReportEvent micDone, "Success", "Application is ready to use!", "report_image.png"
'	End  If
'End If
'
'On Error Goto 0
'####################################################################################










'####################################################################################
'====================================================================================
'*
' Closes the application using it's process
'*
'====================================================================================
'Reporter.Filter = 3
'SystemUtil.CloseProcessByName "diagnostic.exe"
'Reporter.Filter = 0
'Reporter.ReportEvent micDone, "Success", "The application is closed!"

'####################################################################################










'####################################################################################
'====================================================================================
'*
' Recognition of back button using Exist() method.
'*
'====================================================================================

'wait 5
'If Browser("Home").Exist(10) Then
'      	If 	Browser("Home").Page("Home").WebButton("BACK_BTN").Exist(10) Then
'		   	Browser("Home").Page("Home").WebButton("BACK_BTN").Click
'		   	Desktop.CaptureBitmap "report_image.png", True
'		   	Reporter.ReportEvent micDone, "Success", "Ready!", "report_image.png"
'		   'Wait 10
'		   	Desktop.CaptureBitmap "report_image.png", True
'		   	Reporter.ReportEvent micDone, "Success", "Ready!", "report_image.png"
'		Else 
'			Desktop.CaptureBitmap "report_image.png", True
'			Reporter.ReportEvent micFail, "Failled", "Waiting!!", "report_image.png"
'		End  If
'Else 
'	Desktop.CaptureBitmap "report_image.png", True
'	Reporter.ReportEvent micFail, "Failled", "Waiting!!", "report_image.png"
'End If
'
'Wait 10
'####################################################################################









'####################################################################################
'====================================================================================
'*
' In this case, if back_button isn't visible, an error message will be added in report by user. Otherwise, if the button is clicked, will be added a success message.
'*
'====================================================================================

'On Error Resume Next
'Reporter.Filter = 3
'Browser("Home").Page("Home").WebButton("BACK_BTN").Click
'If Err.Number <> 0 Then
'	Reporter.Filter = 0
'	Desktop.CaptureBitmap "report_image.png", True
'	Reporter.ReportEvent micFail, "BACK_BTN", "ErrDescription: " & Err.Description & vbcrlf & "Number error:"  & Err.Number, "report_image.png"
'	
'	Set fso = CreateObject("Scripting.FileSystemObject")
'	Set f = fso.OpenTextFile("C:\log2.txt", 8, True)
'	f.WriteLine "Error: " & Err.Number
'	f.WriteLine "Source: "& Err.Source
'	f.WriteLine "Description: "& Err.Description
'	f.Close
'	Set fso = Nothing
'	Set f = Nothing
'	Err.Clear' Clear the Error
'Else
'	Reporter.Filter = 0
'	Desktop.CaptureBitmap "C:\report_image.png", True
'	Reporter.ReportEvent micDone, "BACK_BTN", "Back button was clicked", "C:\report_image.png"
'End If
'On Error Goto 0' Don't resume on Error
'####################################################################################
