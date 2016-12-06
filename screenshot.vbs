
Option Explicit 

' ------- Declare the variables  ----------------- 
Dim WshShell 


' ------- Blocks of code for the test steps  ------------- 
Sub CaptureScreen 
 	Set WshShell = WScript.CreateObject("WScript.Shell") 
 	WshShell.Run "mspaint" 
	WScript.Sleep 5000 
  
' Activate the IE window so we screen shot it and not something else
 	WShShell.AppActivate "Google"
 	WScript.Sleep 1000   

' Take a screen shot of just the active IE window using [ALT] + [PrtScn]
 	Set Wshshell = CreateObject("Word.Basic") 
 	WshShell.SendKeys "(%{1068})"
 	WScript.Sleep 1000 

' Make Paint active and save the image 
 	WshShell.AppActivate "Untitled - Paint" 
 	WScript.Sleep 1500 
  
	WshShell.sendkeys "^(v)"
 	WScript.Sleep 1500
  
	WshShell.sendkeys "^(s)"
 	WScript.Sleep 1500
   
 	WshShell.sendkeys "testing.jpg"
 	WScript.Sleep 1500
   
 	WshShell.sendkeys "%(s)"
 	WScript.Sleep 1500
End Sub


'----------------------------------

Sub ClosePaintAndIE 
	WshShell.AppClose "Paint" 
 	WScript.Sleep 1500 
	WshShell.AppClose "Google"
 	WScript.Sleep 1500 
End Sub 
 
