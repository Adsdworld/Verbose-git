Option Explicit


' Affiche une boite de dialogue avec un bouton OK
msgbox("USB detector 2000 ? Ok pour continuer")


' Initialisation des variables 
Dim objFSO, objSink, objWMI, hostName, courant_dir
hostName = "."
On Error Resume Next
   Set objFSO     = CreateObject("Scripting.FileSystemObject")
   Set objSink    = WScript.CreateObject("WbemScripting.SWbemSink", "Sink_")
   Set objWMI     = GetObject("winmgmts:\\" & hostName & "\root\cimv2")
   If Err.Number <> 0 Then
      Wscript.Quit
   End If
   objWMI.ExecNotificationQueryAsync objSink, "Select * From __InstanceCreationEvent Within 1 Where " & _
                                             "TargetInstance Isa 'Win32_DiskDrive' And TargetInstance.InterfaceType = 'USB'"
On Error Goto 0


Do
   WScript.Sleep 1
Loop


Sub Sink_OnObjectReady(objEvent, objContext)
   ' Affiche une boite de dialogue avec un bouton OK
   MsgBox("Clé usb detecté")


   On Error Resume Next
      If Err.Number <> 0 Then
         Wscript.Quit
      End If
   On Error Goto 0
End Sub