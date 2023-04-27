Option Explicit


msgbox("pret ? Ok pour continuer")

Dim objFSO, objSink, objWMI, scriptBaseName, hostName, courant_dir
hostName = "."
On Error Resume Next
   Set objFSO     = CreateObject("Scripting.FileSystemObject")
   Set objWMI     = GetObject("winmgmts:\\" & hostName & "\root\cimv2")
   courant_dir = WScript.ScriptFullName 
   scriptBaseName = objFSO.GetBaseName(Wscript.ScriptFullName)
   If Err.Number <> 0 Then
      Wscript.Quit
   End If

On Error Goto 0
   
courant_dir=left(courant_dir,InStrRev(courant_dir,"\"))

MsgBox("J'ai quelque chose a te dire, OK pour continuer")
Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")
' Play audio
oPlayer.URL = courant_dir+"usb.mp3"
oPlayer.controls.play 
While oPlayer.playState <> 1 ' 1 = Stopped
   WScript.Sleep 100
Wend
' Release the audio file
oPlayer.close
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
MsgBox(":\ ADIEU :/")
objShell.Run "C:\WINDOWS\system32\shutdown.exe -l"



On Error Resume Next
   If Err.Number <> 0 Then
      Wscript.Quit
   End If
On Error Goto 0