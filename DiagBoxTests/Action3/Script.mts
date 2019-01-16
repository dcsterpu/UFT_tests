'################################################################################### @@ hightlight id_;_65924_;_script infofile_;_ZIP::ssf50.xml_;_
'====================================================================================
'*
'	Function for pop-ups 
'*
'====================================================================================

Public Function Popups(strBrowser, strPage)
	mytext = Browser(strBrowser).Page(strPage).GetRoProperty("innertext")
	'wait 10
	
	Dim aux
	aux = true
	
	While aux
		If InStr(1, mytext, "AUTHENTICATION") OR InStr(1, mytext, "Do you search for new updates?") OR InStr(1, mytext, "It is necessary to activate the applications in order to be able to use them, you will now activate them") Then
			'UPDATES POP-UP
			If InStr(1, mytext, "Do you search for new updates?") <> 0 Then '?????
				Browser(strBrowser).Page(strPage).WebButton("CANCEL_BTN").Click
				mytext = Browser(strBrowser).Page(strPage).GetRoProperty("innertext")
			End If
			
			'ACTIVATION POP-UP
			If InStr(1, mytext, "It is necessary to activate the applications in order to be able to use them, you will now activate them") <> 0 Then
				Browser(strBrowser).Page(strPage).WebButton("CANCEL_BTN").Click
				mytext = Browser(strBrowser).Page(strPage).GetRoProperty("innertext")
				Call CloseApp()
			End If
			
			'AUTHENTICATION POP-UP
			If InStr(1, mytext, "AUTHENTICATION") <> 0 Then 
				If InStr(1, mytext, "Select the marque") Then
					'Select the marque??
					'Browser(strBrowser).Page(strPage).WebButton("select").Click
			   	 	'Browser(strBrowser).Page(strPage).WebButton("select").ChildObjects "DS"
			   	 	Browser(strBrowser).Page(strPage).WebEdit("UserName").Set "E518720" @@ hightlight id_;_10000000_;_script infofile_;_ZIP::ssf38.xml_;_
			    	Browser(strBrowser).Page(strPage).WebEdit("UserPass").Set "Cst78788"
			    	Browser(strBrowser).Page(strPage).WebButton("OK_BTN").Click
					wait 2
				Else
					Browser(strBrowser).Page(strPage).WebEdit("UserName").Set "E518720"
					Browser(strBrowser).Page(strPage).WebEdit("UserPass").Set "Cst78788"
					Browser(strBrowser).Page(strPage).WebButton("OK_BTN").Click
					wait 2
				End If
				
				mytext = Browser(strBrowser).Page(strPage).GetRoProperty("innertext")
				
				While InStr(1, mytext, "Please wait")
					mytext = Browser(strBrowser).Page(strPage).GetRoProperty("innertext")
					If InStr(1, mytext, "AUTHENTICATION ERROR") <> 0 Then
						Reporter.Filter = 0
						Desktop.CaptureBitmap "report_image.png", True
						Reporter.ReportEvent micFail, "AUTHENTICATION ERROR", "Incorrect identifier", "report_image.png"
						Browser(strBrowser).Page(strPage).WebButton("CANCEL_BTN").Click
					End If
				Wend		
			Else
				Reporter.Filter = 0
				Desktop.CaptureBitmap "report_image.png", True
				Reporter.ReportEvent micWarn, "No AUTHENTICATION pop-up", "No AUTHENTICATION pop-up", "report_image.png"
			End If
			
		Else
			aux = false
		End If	
	Wend
	
	' for other pop-ups
	If Browser(strBrowser).Page(strPage).WebButton("CANCEL_BTN").Exist(5) Then
		Desktop.CaptureBitmap "report_image.png", True
		Reporter.ReportEvent micWarn, "Pop-up", "New pop-up", "report_image.png"
		Browser(strBrowser).Page(strPage).WebButton("CANCEL_BTN").Click
	End  If
	
End  Function
'
'Call Popups("Home","Home")
'####################################################################################








'################################################################################### @@ hightlight id_;_65924_;_script infofile_;_ZIP::ssf50.xml_;_
'====================================================================================
'*
'	Version of application
'*
'====================================================================================

Public Function VersionOfApp(text, versionDb)
   Dim regExVersion, Match, Matches   ' Create variable.
   Set regExVersion = New RegExp   ' Create a regular expression.
   regExVersion.Pattern = "v[0-9][0-9]\.[0-9][0-9]"   ' Set pattern
   Set Matches = regExVersion.Execute(text)   ' Execute search.
   For Each Match in Matches   ' Iterate Matches collection.
      currentVersionDb = Match.Value
   Next
   If currentVersionDb = versionDb Then
	    Desktop.CaptureBitmap "report_image.png", True
	    Reporter.ReportEvent micDone, "Version of application", "The version of DiagBox is "  & currentVersionDb, "report_image.png"
   Else
		Desktop.CaptureBitmap "report_image.png", True
	    Reporter.ReportEvent micFail, "Version of application", "The version of DiagBox is "  & currentVersionDb, "report_image.png"
	    wait 10
	    Call CloseApp()
   End If
End Function

'version = "v09.38"
'mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'Call VersionOfApp(mytext, version)
'###################################################################################








'################################################################################### @@ hightlight id_;_65924_;_script infofile_;_ZIP::ssf50.xml_;_
'====================================================================================
'*
' Open application
' Verify if a pop-up appeared
' Verify your version of application
'*
'====================================================================================
'SystemUtil.Run "C://AWRoot/bin/launcher/LctPOLUX.exe" 
'wait 25
'mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'While InStr(1, mytext, "Please wait") <> 0 
'	mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'Wend
'wait 10
'Call Popups("Home", "Home")
'version = "v09.37"
'mytext = Browser("Home").Page("Home").GetRoProperty("innertext")
'Call VersionOfApp(mytext, version)
'####################################################################################









'#################################################################################### @@ hightlight id_;_10000000_;_script infofile_;_ZIP::ssf32.xml_;_
'====================================================================================
'*
' Authentication
'*
'==================================================================================
Public Function Authentication(strBrowser, strPage)
	Reporter.Filter = 3
	Browser(strBrowser).Page(strPage).WebButton("jeton_user").Click
	wait 2
	While InStr(1, mytext, "AUTHENTICATION") = 0
			mytext = Browser(strBrowser).Page(strPage).GetRoProperty("innertext")
	Wend
	Call Popups(strBrowser, strPage)
End Function

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
'   Reporter.ReportEvent micFail, "Failled", "The path is invalid!", "report_image.png"
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
' The application is closed using its process
'*
'====================================================================================
Public Function CloseApp()
	Reporter.Filter = 3
	SystemUtil.CloseProcessByName "diagnostic.exe"
	Reporter.Filter = 0
	Reporter.ReportEvent micDone, "Success", "The application is closed!"
End  Function
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
