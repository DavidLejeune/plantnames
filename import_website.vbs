


 outFile="list_of_plants.txt"
 outPlant="plant.txt"
 Set objFSO=CreateObject("Scripting.FileSystemObject")
 Const Unicode = -1
 With (CreateObject("Scripting.FileSystemObject"))
       If .FileExists(outFile) Then
         .DeleteFile(outFile)
       Else

       End If
End With
Set objFile = objFSO.CreateTextFile(outFile,2,True)


url = "http://www.megabytedata.com/MB054/getlist.asp?lang=ER&langname=Esperanto"
Set ie = CreateObject("InternetExplorer.Application")
ie.Navigate url

While ie.ReadyState <> 4
  WScript.Sleep 100
Wend

dim table
set table = ie.document.getElementById( "example" )
dim cellContent0() 'link + id
dim cellContent1() 'Esperanto name
dim cellContent2() 'Latin name
redim cellContent0( table.rows.length - 1 )
redim cellContent1( table.rows.length - 1 )
redim cellContent2( table.rows.length - 1 )


dim table2


'this will loop all the rows of the table and pull all the column 0 and column 1
'for i = 0 to table.rows.length - 1
for i = 0 to 10
 cellContent0(i) = table.rows(i).cells(0).innerhtml
 strID = Replace(cellContent0(i),"<a href=" & """" & "allnames.asp?ID=","")
 strID = Replace(strID,"""" & ">All names</a>","")
 cellContent1(i) = table.rows(i).cells(1).innerText
 cellContent2(i) = table.rows(i).cells(2).innerText

'Wscript.echo strID

  if isNumeric(strID) Then

   objFile.Write strID & "," & cellContent1(i) & "," & cellContent2(i) & vbCrLf

   outPlant="plant.txt"
   'wscript.echo "strID : " & strID
    With (CreateObject("Scripting.FileSystemObject"))
      If .FileExists(outPlant) Then
        .DeleteFile(outplant)
      Else

      End If
    End With

   Set objPlant = objFSO.CreateTextFile(outPlant,True , True)
   url2 = "http://www.megabytedata.com/MB054/allnames.asp?ID=" & strID
   strText = ""

    With CreateObject("MSXML2.XMLHTTP")
      .open "GET", url2 , False
      .send
      'WScript.Echo .responseText
      strText = .responseText
      'WScript.Echo strText
      'Wscript.echo len(strText)
    End With

    objPlant.Write strText & "" &  vbCrlf
    objPlant.Close
    wscript.sleep 1

   wscript.echo strID & "," & cellContent1(i) & "," & cellContent2(i) ''& "," & PagecellContent0(i) & "," & PagecellContent1(i) & "," & PagecellContent2(i)
  End If

next

objFile.close
















' filename = outPlant
'
' Set fso = CreateObject("Scripting.FileSystemObject")
' Set f = fso.OpenTextFile(filename,1)
'
' Do Until f.AtEndOfStream
'   strLine = f.ReadLine
'   WScript.Echo f.ReadLine
' Loop
'
' f.Close
'


'
' strSearchString = "ABCDEFGHIJK, NDPSGW PORT=LPR HOSTNAME=R2333_HP_1100 ABCDEFGHIJK"
'
' intStart = InStr(strSearchString, "HOSTNAME=")
'
' intStart = intStart + 9
'
'
' strText = Mid(strSearchString, intStart, 250)
'
'
' For i = 1 to Len(strText)
'
'     If Mid(strText, i, 1) = " " Then
'
'         Exit For
'
'     Else
'
'         strData = strData & Mid(strText, i, 1)
'
'     End If
'
' Next
'
'
' Wscript.Echo strData
'


























' url2 = "http://www.megabytedata.com/MB054/allnames.asp?ID=" & strID
'   Set ie2 = CreateObject("InternetExplorer.Application")
'  ie2.Navigate url2
'
'  While ie2.ReadyState <> 4
'    WScript.Sleep 100
'  Wend
'
'  set table2 = ie2.document.getElementsByTagName( "table" ).Item(2)
'  dim PagecellContent0() 'link + id
'  dim PagecellContent1() 'Esperanto name
'  dim PagecellContent2() 'Latin name
'  redim PagecellContent0( table2.rows.length - 1 )
'  redim PagecellContent1( table2.rows.length - 1 )
'  redim PagecellContent2( table2.rows.length - 1 )
'
'
'  PagecellContent0(i) = table2.rows(i).cells(0).innerText
'  PagecellContent1(i) = table2.rows(i).cells(1).innerText
'  PagecellContent2(i) = table2.rows(i).cells(2).innerText


' theURL = "www.megabytedata.com/MB054/allnames.asp?ID=" & strID
' Set ie2 = CreateObject("InternetExplorer.Application")
' with ie2
'   .Navigate("http://" & theURL)
'   Do until .ReadyState = 4
'      WScript.Sleep 50
'   Loop
'   With .document
'     set theTables = .all.tags("table")
'     nTables = theTables.length
'     for each table in theTables
'       s = s & table.rows(0).cells(0).innerText _
'         & vbNewLine & vbNewLine
'     next
'     wsh.echo "Number of tables:", nTables, vbNewline
'     wsh.echo "First table first cell:", s
'     ' get the data with an ID
'     wscript.echo ie2.document.getelementbyid("td").innerHtml
'   End With
' End With




' filename = "D:\DaLe\David\Projects\PlantNames\table.txt"
'
' Set fso = CreateObject("Scripting.FileSystemObject")
' Set f = fso.OpenTextFile(filename)
'
' count=0
' Do Until f.AtEndOfStream
'   count=count+1
'   WScript.Echo f.ReadLine
'   strLine = f.ReadLine
'
'   outputArray = split(strLine,"</TR>")
'
'   for each x in outputArray
'     message = message & x & vbCRLF
'   next
'
'   echo message
'
'
'
' Loop
'
' f.Close


' url = "http://www.megabytedata.com/MB054/getlist.asp?lang=ER&langname=Esperanto"
'
' Set ie = CreateObject("InternetExplorer.Application")
' ie.Navigate url
'
' While ie.ReadyState <> 4
'   WScript.Sleep 100
' Wend
'
' ' 'get 3rd iframe in page
' ' Set iframe = ie.document.getElementsByTagName("iframe").Item(2).contentWindow
' ' 'get 1st table in iframe
' Set tbl = ie.document.getElementById("example")
' 'get 4th cell in table
' Set td  = tbl.getElementsByTagName("tr").Item(1)
'
' Wscript.echo td.innerText
