Option Explicit

' Bien placer le vbs et le fichier dans le même dossier
Dim filename
filename = "au revoir"
' Testé sur du mp3 et du mp4, cela doit aussi fonctionner sur d'autres format
Dim extention
extention = "."+"mp3"


' Affiche une boite de dialogue avec un bouton OK
msgbox("pret ? Ok pour continuer")


' Initialisation des variables 
Dim courant_dir
On Error Resume Next
   courant_dir = WScript.ScriptFullName 
   If Err.Number <> 0 Then
      Wscript.Quit
   End If
On Error Goto 0


' Construit une chaine de caractères du chemin du dossier
courant_dir=left(courant_dir,InStrRev(courant_dir,"\")) 


' Affiche une boite de dialogue avec un bouton OK
MsgBox("J'ai quelque chose a te dire, OK pour continuer")


' Définir variable
Dim oPlayer
' Affecter la variable à CreateObject("WMPlayer.OCX")
Set oPlayer = CreateObject("WMPlayer.OCX")
' Ouvre un audio invisible 
oPlayer.URL = courant_dir+filename+extention
oPlayer.controls.play 
' Attends tant que la lecture n'est pas arrêter state 1 = pause/stop
While oPlayer.playState <> 1 ' 
   WScript.Sleep 100
Wend
' Ferme le fichier audio
oPlayer.close


' Définir une variable
Dim objShell
' Affecter la variable à WScript.CreateObject("WScript.Shell")
Set objShell = WScript.CreateObject("WScript.Shell")
' Affiche une boite de dialogue avec un bouton OK
MsgBox(":\ ADIEU :/")
' Execute la commande C:\WINDOWS\system32\shutdown.exe -l dans l'objet/terminal
objShell.Run "C:\WINDOWS\system32\shutdown.exe -l"


On Error Resume Next
   If Err.Number <> 0 Then
      Wscript.Quit
   End If
On Error Goto 0