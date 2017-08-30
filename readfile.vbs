outPlant="D:\DaLe\David\Projects\PlantNames\plant.txt"
filename = "plant.txt"


Const ForReading = 1

Const TriStateTrue = -1

' Set objFSO = CreateObject("Scripting.FileSystemObject")
'
'
' Set objFile = objFSO.OpenTextFile(filename, ForReading,False,TriStateTrue)
'
' Do Until objFile.AtEndOfStream
'   WScript.Echo objFile.ReadLine
' Loop
' 'strText = objFile.ReadAll
'
' objFile.Close


'Wscript.Echo strText






' Set f = fso.OpenTextFile(filename)
'
' Do Until f.AtEndOfStream
'   WScript.Echo f.ReadLine
' Loop
'
' f.Close


Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename, ForReading,False,TriStateTrue)

count=0
strTranslated=""
sBoolLang = False

  strResult =""
Do Until f.AtEndOfStream
  count=count+1
  strLine = f.ReadLine
  strLangIntro = "<TD width=" & """"  & "200" & """" & "><font color=" & """" & "#800000" & """" & "><b>"
  strLangIntro2 = "<TD width=" & """" & "200" & """" & " valign=" & """" & "top" & """" & " align=" & """" & "left" & """" & "><font color=" & """" & "#800000" &  """" & "><b>"
  strLangOutro = " : </b></font></TD>"
  strLangOutro2 = " :</b> </font></TD>"

  strTranIntro = "<TD><font color=" & """" & "#000000" & """" & ">"
  strTranIntro2 = "<TD class=multil>"
  strTranOutro = " </font></TD>        </TR>"
  strTranOutro2 = "&nbsp;&nbsp;                          </font></TD>        </TR>"


  If sBoolLang = True Then
    strTranslated = strTranslated & strLine
    If instr(strLine , "</TR>") Then
    strTranslated = Replace (strTranslated, strTranIntro , "")
    strTranslated = Replace (strTranslated, strTranIntro2 , "")
    strTranslated = Replace (strTranslated, strTranOutro , "")
    strTranslated = Replace (strTranslated, strTranOutro2 , "")
      Wscript.echo LTrim(strTranslated)
      sBoolLang = False
      strResult = strResult & "," & LTrim(strTranslated)
      strTranslated=""


    End If
  End If



  if InStr(strLine, strLangIntro) > 0 or InStr(strLine, strLangIntro2) > 0 Then
  strLine = Replace (strLine, strLangIntro , "")
  strLine = Replace (strLine, strLangIntro2 , "")
  strLine = Replace (strLine, strLangOutro , "")
  strLine = Replace (strLine, strLangOutro2 , "")
    'WScript.Echo count & " : " & LTrim(strLine)
    sBoolLang = True
  end if
Loop
strResult = Replace (strResult, "&nbsp;" , "")
strResult = Replace (strResult, "  " , "")
Wscript.echo strResult

f.Close
