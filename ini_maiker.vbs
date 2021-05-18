Dim objFS, ttFil, objFSs,ttFile
Set objFS = Createobject("Scripting.FileSystemObject")
Set objFSs = Createobject("Scripting.FileSystemObject")
Set ttFile = objFSs.OpenTextFile("C:\work\vbs\города.txt",1)

Do  While Not ttFile.AtEndOfStream
  Sx=ttFile.Readline
Set ttFil = objFS.OpenTextFile("C:\work\vbs\ini.txt",1)
'msgbox SX
Do  While Not ttFil.AtEndOfStream
  Sxx=ttFil.Readline
  Sxx= Sx & "%20" & Sxx
  WriteTXT Sxx
'   msgbox Sxx
 loop
ttFil.close
i = i+1
'msgbox i
loop
ttFile.close

Sub WriteTXT(LogMessage)
Const ForAppending = 8
dim A, B, C, objLogFile, objFSs
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C="0" & C end if
Set objFSs = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSs.OpenTextFile("C:\work\vbs\ini1.txt" , ForAppending, TRUE)
objLogFile.WriteLine(LogMessage)
End Sub