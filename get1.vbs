'адрес страницы в инете 
strURL = "http://maps.yandex.ru/?text=%D0%92%D0%BE%D0%B5%D0%BD%D0%BA%D0%BE%D0%BC%D0%B0%D1%82%D1%8B%2C%20%D0%BA%D0%BE%D0%BC%D0%B8%D1%81%D1%81%D0%B0%D1%80%D0%B8%D0%B0%D1%82%D1%8B&results=20000&l=map"

'сохранить как... 
strFile = "C:\work\work\index.html" 

strText = TextFromHTML( strURL ) 

Set fso = CreateObject("Scripting.FileSystemObject") 
Set fileHandle = fso.CreateTextFile( strFile, True ) 
    fileHandle.Write strText 
    fileHandle.Close 

'функция выдирания чистого текста со страницы, без html-тегов 
Function TextFromHTML( URL ) 
  Set objIE = CreateObject( "InternetExplorer.Application" ) 
      objIE.Navigate URL 
  Do Until objIE.ReadyState = 4 
      WScript.Sleep 100 
  Loop 
  TextFromHTML = objIE.Document.Body.InnerText 
  objIE.Quit 
End Function 

WScript.Quit