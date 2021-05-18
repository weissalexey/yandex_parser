'======================================================================
'  Парсер для яндекса читает запросы с ini1.txt и собирает в выходном файле 
'  ini1.txt ДЕЛАЕТСЯ с помощью ini_maiker.vbs из файлов ini.txt и города.txt
'  Создатель скрипта WEISS
' id.txt добавлен для удаления повторов
'
'
'
'
'======================================================================
WriteLog "========== Запуск скрипта ==========" 
WriteLog "========== getYAndex.vbs ==========="

'
' localhost
' shop_user
' K67Th7Z6xFPxFAxb
' shop
'
'
Const scriptVer  = "1.0"
Const DownloadType = "binary"
dim strURL
Dim filespec
Dim objFSO, txtFile, x, t, sw, pr, ppp, DownloadDest, iii
Dim objFSs, ttFile, Sx, pr12
filespec = "C:\work\work\index.html"
Set objFSO = Createobject("Scripting.FileSystemObject")
writelog ("Нчали ")
if (TestWorkingFail("C:\work\vbs\ini1.txt")) then
writelog "проверка"
Set txtFile = objFSO.OpenTextFile("C:\work\vbs\ini1.txt",1)
writelog ("Нашли C:\work\vbs\ini1.txt") 
Do  While Not txtFile.AtEndOfStream
 iii=iii+1
WriteLog "Считали строку номер " & iii
 filespec = "C:\work\work\index" & iii &".html"
  x=txtFile.Readline
  DownloadDest ="http://maps.yandex.ru/?text=" & x &"&results=999999"
 writelog ("Запрос " & DownloadDest) 
 writelog (filespec & "  Проверяем наличие") 
if (TestWorkingFail(filespec)) then
   writelog (filespec & "  Файл уже существует удаляем") 
   DeleteWorkingFail(filespec)   
 End If

If Wscript.Arguments.Named.Exists("h") Then
  Wscript.Echo "Usage: http-download.vbs"
  Wscript.Echo "version " & scriptVer
  WScript.Quit(intOK)
End If

getit(filespec)
if (TestWorkingFail(filespec)) then
writelog ("Файл " & filespec & " обнаружен  парсим" )
Set objFSs = Createobject("Scripting.FileSystemObject")
Set ttFile = objFSs.OpenTextFile(filespec,1)

Do  While Not ttFile.AtEndOfStream
  Sx=ttFile.Readline
  If InStr(Sx, "CompanyMetaData") Then
      t = InStr(Sx, "CompanyMetaData")
      sw = right( Sx , len(Sx) - t -22)
      Do  While t > 0
      t = InStr(sw, "CompanyMetaData")
      pr = left(sw ,t)
      pr = Replace(pr, "name" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr,"names" & chr(34) & ":[" & chr(34) , "")
      pr = Replace(pr,"]" , "")
      pr = Replace(pr,"address" & chr(34) & ":" & chr(34) , "")     '"},
      pr = Replace(pr,"postalCode" & chr(34) & ":" & chr(34) , "")     
      pr = Replace(pr,"AddressDetails" & chr(34) & ":{" & chr(34) & "Country" & chr(34) & ":{" & chr(34) & "AddressLine" & chr(34) & ":"& chr(34), "")     
      pr = Replace(pr,chr(34) & "CountryNameCode" & chr(34) & ":", "")  
      pr = Replace(pr,"CountryName" & chr(34) & ":" & chr(34) , "")   
      pr = Replace(pr,"AdministrativeArea" & chr(34) & ":{" & chr(34) & "AdministrativeAreaName" & chr(34) & ":" & chr(34) , "")    
      pr = Replace(pr,"SubAdministrativeArea" & chr(34) & ":{" & chr(34) & "Locality" & chr(34) & ":{" & chr(34) & "LocalityName" & chr(34) & ":"& chr(34), "")   
      pr = Replace(pr,"Thoroughfare" & chr(34) & ":{" & chr(34) & "ThoroughfareName" & chr(34) & ":" & chr(34) , "")    
      pr = Replace(pr,"Premise" & chr(34) & ":{" & chr(34) & "PremiseNumber" & chr(34) & ":" & chr(34) , "")  
      pr = Replace(pr,"Premise" & chr(34) & ":{" & chr(34) & "PremiseName" & chr(34) & ":" & chr(34) , "")  
      pr = Replace(pr,"}}," , ",")
      pr = Replace(pr,"Thoroughfare" & chr(34) & ":{" & chr(34) & "Premise" & chr(34) & ":" & chr(34) , "")  
      pr = Replace(pr,"SubPremise" & chr(34) & ":{" & chr(34) & "SubPremiseName" & chr(34) & ":" & chr(34) , "")    
      pr = Replace(pr,"PostalCode" & chr(34) & ":{" & chr(34) & "PostalCodeNumber" & chr(34) & ":" & chr(34) , "")    
      pr = Replace(pr,"}}}}}}}}" , "")
      pr = Replace(pr, "}," ,  ",")
      pr = Replace(pr, "url" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr,"urls" & chr(34) & ":[" & chr(34) , "")
      pr = Replace(pr,"Categories" & chr(34) & ":[{" & chr(34) , "") 
      pr = Replace(pr,"InternalCategoryInfo" & chr(34) & ":{" & chr(34) & "AppleData" & chr(34) & ":{" & chr(34) & "acid" & chr(34) & ":"& chr(34), "")    
      pr = Replace(pr, "level" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr,"}}," , ",")
      pr = Replace(pr,",{" , ",")
      pr = Replace(pr,"internal" & chr(34) & ":{" & chr(34) & "id" & chr(34) & ":" & chr(34) , "")    
      pr = Replace(pr, "acid" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr, "level" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr,"}}," , ",")
      pr = Replace(pr,"Phones" & chr(34) & ":[{" & chr(34) & "type" & chr(34) & ":" & chr(34) , "")    
      pr = Replace(pr, "formatted" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr, "country" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr, "prefix" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr, "number" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr, "info" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr,"},{" , ",")
      pr = Replace(pr, "type" & chr(34) & ":" & chr(34), "")
      pr = Replace(pr,"}," , ",")
      pr = Replace(pr,"Hours" & chr(34) & ":{" & chr(34) & "Availability" & chr(34) & ":{" & chr(34) , "")  
      pr = Replace(pr,chr(34) & ":" & chr(34) , "")
      pr = Replace(pr, "," & chr(34) & "text", "")  
      pr = Replace(pr, "Geo" & chr(34) & ":{" & chr(34), "")
      pr = Replace(pr, "InternalCompanyInfo" & chr(34) & ":{" & chr(34), "")      
      pr = Replace(pr, "internal" & chr(34) & ":{" & chr(34), "")   
      pr = Replace(pr, "," & chr(34) & "synonym", "")  
      pr = Replace(pr,"geoid" & chr(34) & ":30," & chr(34) & "company_id" , "")  
      If InStr(pr, "featureData") Then ppp = InStr(pr,"featureData" ) -3 Else ppp = 0  End If
      pr = left( pr ,  ppp  )
      pr12= Right(left (pr,11),10)

      Writelog ("Поверяем Наличие такого Экземпряра в Базе")

      If len( pr12) > 9 then 
      if (TestNombDoble (pr12)) then 
       writeid pr12 
       WriteTXT pr 
      end if
      end if
   
      Sw = right( sw, len(sw) - t -22)
      loop
  End If
Loop
ttFile.Close
DeleteWorkingFail(filespec)  
WriteLog "Разпарсили и убрали за собой"
End If


 loop
dim A, B, C
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C="0" & C end if
A = "C:\work\work\" & A & B & C & ".txt" 
        WriteLog "========== Отчет о Парсере ==========" 
        WriteLog "========== getYAndex.vbs ==========="
	WriteLog "Было задано " & iii & " запросов." 
	WriteLog "Скриптом было обработано - " & iii  & " файлов."
        WriteLog "Выходной файл - " &  A 
	WriteLog "========== Выполнение скрипта завершено ==========" 
	

End If

Sub WriteLog(LogMessage)
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("C:\work\log\" & A & B & C & ".log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "] " & LogMessage)
End Sub

Sub WriteTXT(LogMessage)
Const ForAppending = 8
dim A, B, C, objLogFile, objFSs
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C="0" & C end if
Set objFSs = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSs.OpenTextFile("C:\work\work\YABASE.txt" , ForAppending, TRUE)
objLogFile.WriteLine(LogMessage)
End Sub

    function getit(LocalFile)
    writelog ("Качаем данные и сохраняем на диск") 
    dim xmlhttp
    set xmlhttp=createobject("MSXML2.XMLHTTP.3.0")
    strURL = DownloadDest
    xmlhttp.Open "GET", strURL, false
    xmlhttp.Send
    If xmlhttp.Status = 200 Then
    writelog ("Соединение установленно данные получены") 
    Dim objStream, ress
    set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
    objStream.Write xmlhttp.responseBody
    objStream.SaveToFile LocalFile
    objStream.Close
    set objStream = Nothing
    writelog ("Все сохранено в " & LocalFile) 
    Else
    writelog ("*************************Проверь соединение или запрос*************************") 
    End If
    set xmlhttp=Nothing
End function 


function verif (id, pr)
   Const ForAppending = 8
   dim A, B, C, F, objLogFile, objFSs
   Set objFSs = CreateObject("Scripting.FileSystemObject")
   Set objLogFile = objFSs.OpenTextFile("C:\work\vbs\id.txt" ,1)
   C = false
   Do  While Not objLogFile.AtEndOfStream
   A   = objLogFile.readline
   If len(id) > 9 then

   If A = id then C = true  End If
   end if   
   loop
   if c = false then
   If len(id) > 9 then    
   WriteTXT pr
   set B  = CreateObject("Scripting.FileSystemObject")
   Set F= B.OpenTextFile("C:\work\vbs\id.txt" , ForAppending, TRUE)
    F.WriteLine(id)
    F.close
    end if
   else
   Writelog ("Такая запись уже существует ")
   end if
   End function 

function TestNombDoble(fnom)
                                        filename = "C:\work\vbs\id.txt"
                                        INPUT_FILE_1 = (filename)
					Const ForReading = 1
					Dim dicData1, strKey, fso
					Dim filInput1, strLine
					Set fso = CreateObject("Scripting.FileSystemObject")
					Set filInput1 = fso.OpenTextFile(INPUT_FILE_1, ForReading)
					Set dicData1 = CreateObject("Scripting.Dictionary")
                                        
					While Not filInput1.AtEndOfStream
                                        i=i+1
					strLine = filInput1.ReadLine
					dicData1.Add (i) , (strLine) 
					Wend
					filInput1.Close
                                            TestNombDoble = Falce
                                      For Each strKey In dicData1
                                             If dicData1.Item(strKey) <> fnom Then
                                             TestNombDoble =True
                                             End if
                                      Next
end function

function TestWorkingFail(fname)
	Dim fso, ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(fname) then
	  TestWorkingFail = true
	else
	 TestWorkingFail = false
	end if
end function

function DeleteWorkingFail(fname)
	Dim fso, ts
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile (fname)
end function

sub writeid (id)
Const ForAppending = 8
dim  B, F
set B  = CreateObject("Scripting.FileSystemObject")
   Set F= B.OpenTextFile("C:\work\vbs\id.txt" , ForAppending, TRUE)
    F.WriteLine(id)
    F.close
end sub