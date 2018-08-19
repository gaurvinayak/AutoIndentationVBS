
Filepath=BrowseForFile()
ArrFileName=Split(Filepath,"\")
FileName=ArrFileName(UBound(ArrFileName))
ArrFileName=Split(FileName,".")
strFileName="Temp-"&ArrFileName(0)&".txt"

For i=0 to UBound(ArrFileName) 

If i=UBound(ArrFileName) Then
UpdateFileName=UpdateFileName&"\"&strFileName
Exit for
elseif i=0 thenUpdateFIleName=ArrFileName(i)
Else
UpdateFileName=UpdateFileName&"\"&ArrFileName(i)
End if

Next

Function fnWait(ntime)
sTime=Now()
Do while DateDiff("s",sTime,Now())<ntime
loop
End Function


Function copyConetent(Filepath,UpdateFileName)
ArrFileName1=Split(Filepath,"\")
FileName=ArrFileName1(UBound(ArrFileName1))
Set objFileToWrite= CreateObject("Scripting.FileSystemObject").OpenTextFile(UpdateFileName,2,True)
Set objFileToWrite= Nothing
Dim oShell
Set oShell = CreateObject("WScript.Shell")
oShell.run "notepad.exe "&UpdateFileName ,3
oShell.SendKeys "^"
call fnWait(3)
oShell.run "notepad.exe "&Filepath ,3
call fnWait(6)
oShell.SendKeys "^a"
oShell.SendKeys "^c"
call fnWait(3)
oShell.run "notepad.exe "&UpdateFileName ,3
call fnWait(3)
oShell.SendKeys "^a"
call fnWait(1)
oShell.SendKeys "^v"
oShell.SendKeys "^s"
call fnWait(3)
oShell.run "taskkill /im notepad.exe", , True
set oShell= Nothing
End Function



Call copyConetent(Filepath,UpdateFileName)


strTime=Now()

If UpdateFileName<>"" Then
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(UpdateFileName,1)
strFileText=objFileToRead.ReadAll()
objFileToRead.Close
Set objFileToRead= Nothing
x=BeautifyVBS(strFileText,4)
Set objFileToWrite= CreateObject("Scripting.FileSystemObject").OpenTextFile(UpdateFileName,2,True)
objFileToWrite.WriteLine(x)
objFileToWrite.Close
Set objFileToWrite=Nothing

' Call copyConetent(UpdateFileName,Filepath)
EndTime=Now()

Msgbox "Indentation Completed !"
Else
Msgbox "Please Correct your File Path"
End If



Function BrowseForFile()
    With CreateObject("WScript.Shell")
        Dim fso: set fso =CreateObject("Scripting.FileSystemObject")
        Dim tempFolder : Set tempFolder =fso.getTempName() &".hta"
        Dim path : path = "HKCU\Volatile Environment\MsgResp"
        With tempFolder.CreateTextFile(TempName)
            .Write "<input type=file name=f>" &_
            "<script>f.click();(new ActiveXObject('WScript.Shell'))"&_
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.Value);"&_
            "close();</Script>"
            .Close
        End With
        .Run tempFolder &"\"& tempName,1,True
        BrowseForFile=.regRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder &"\"& tempName
    End With
End Function





Function BeautifyVBS (sSource, nTabSpacing)
' Takes VBScript source code and rebuilds the indentation.
Dim sRawLine, sLine, sTest, iIndentIndex, iIndex, oS, sWhiteSpace, bAdjustIndent, bInQuote, aRows
 
Dim aKey(34)
 
Const INDENT_String = 0
Const INDENT_Exeception_String = 1
Const INDENT_Pre_Indent = 2
Const INDENT_Post_Indent = 3
 
' The indent and unindent list is as complete as I could make it (from the MS VBScript reference).
aKey (0) = ARRAY ("if ", " then", 0, 1)
aKey (1) = ARRAY ("select ", "", 0, 2)
aKey (2) = ARRAY ("sub ", "", 0, 1)
aKey (3) = ARRAY ("function ", "", 0, 1)
aKey (4) = ARRAY ("do ","", 0, 1)
aKey (5) = ARRAY ("while ","", 0, 1)
aKey (6) = ARRAY ("for ","", 0, 1)
aKey (7) = ARRAY ("case ", "", -1, 1)
aKey (8) = ARRAY ("with ","", 0, 1)
aKey (9) = ARRAY ("class ","", 0, 1)
aKey (10) = ARRAY ("public sub","", 0, 1)
aKey (11) = ARRAY ("private sub ","", 0, 1)
aKey (12) = ARRAY ("public function ","", 0, 1)
aKey (13) = ARRAY ("private function ","", 0, 1)
aKey (14) = ARRAY ("property get ","", 0, 1)
aKey (15) = ARRAY ("public property get ","", 0, 1)
aKey (16) = ARRAY ("private property get ","", 0, 1)
aKey (17) = ARRAY ("property let ","", 0, 1)
aKey (18) = ARRAY ("public property let ","", 0, 1)
aKey (19) = ARRAY ("private property let ","", 0, 1)
aKey (20) = ARRAY ("property set ","", 0, 1)
aKey (21) = ARRAY ("public property set ","", 0, 1)
aKey (22) = ARRAY ("private property set ","", 0, 1)
aKey (23) = ARRAY ("else ","", -1, 1)
aKey (24) = ARRAY ("elseif ","", -1, 1)
aKey (25) = ARRAY ("end if", "", -1, 0)
aKey (26) = ARRAY ("end select", "", -2, 0)
aKey (27) = ARRAY ("end sub", "", -1, 0)
aKey (28) = ARRAY ("end function", "", -1, 0)
aKey (29) = ARRAY ("loop", "", -1, 0)
aKey (30) = ARRAY ("wend", "", -1, 0)
aKey (31) = ARRAY ("next", "", -1, 0)
aKey (32) = ARRAY ("end class", "", -1, 0)
aKey (33) = ARRAY ("end property", "", -1, 0)
aKey (34) = ARRAY ("end with", "", -1, 0)
 
sWhiteSpace = " " & vbTab
 
Set oS = CreateObject("ADODB.Stream")
oS.Type = 2   ' ASCII
oS.Open
 
iIndentIndex = 0
For Each sRawLine in Split (sSource, vbCrLf)
 
' Remove all whitespace on the left
iIndex = 1
If Len (sRawLine) > 0 Then
Do While iIndex <= Len (sRawLine)
If Instr (sWhiteSpace, Mid (sRawLine, iIndex, 1)) = 0 Then
Exit Do
End If
iIndex = iIndex + 1
Loop
End If
If iIndex > Len (sRawLine) Then
sLine = ""
Else
sLine = Mid (sRawLine, iIndex)
End If
 
' Remove all whitespace on the right
iIndex = Len (sLine)
Do While iIndex > 0
If Instr (sWhiteSpace, Mid (sLine, iIndex, 1)) = 0 Then
Exit Do
End If
iIndex = iIndex - 1
Loop
If iIndex < Len (sLine) Then
sLine = Left (sLine, iIndex)
End If
 
sTest = LCase (LTrim (sLine))
' Find any in-line comment marker, and truncate the comment if it exists.
bInQuote = False
For iIndex = 1 To Len (sTest)
If Not bInQuote And Mid (sTest, iIndex, 1) = "'" Then
Exit For
End If
If Mid (sTest, iIndex, 1) = """" Then
bInQuote = Not bInQuote
End If
Next
If iIndex < Len (sTest) Then  ' Truncate comment
sTest = Left (sTest, iIndex - 1)
' Truncate whitespace again
iIndex = Len (sTest)
Do While iIndex > 0
If Instr (sWhiteSpace, Mid (sTest, iIndex, 1)) = 0 Then
Exit Do
End If
iIndex = iIndex - 1
Loop
If iIndex < Len (sTest) Then
sTest = Left (sTest, iIndex)
End If
End If
 
sTest = LCase (LTrim (sTest)) & SPACE (32)
 
' Adjust Indentation as needed
bAdjustIndent = False
For iIndex = 0 To UBound (aKey, 1)
If Left (sTest, LEN (aKey(iIndex)(INDENT_String))) = aKey(iIndex)(INDENT_String) Then
If LEN(aKey(iIndex)(INDENT_Exeception_String)) = 0 Or Right (RTrim (sTest), LEN (aKey(iIndex)(INDENT_Exeception_String))) = aKey(iIndex)(INDENT_Exeception_String) Then
bAdjustIndent = True
Exit For
End If
End If
Next
 
If bAdjustIndent Then
iIndentIndex = iIndentIndex + aKey(iIndex)(INDENT_Pre_Indent)
If iIndentIndex < 0 Then iIndentIndex = 0
End If
 
If nTabSpacing <= 0 Then
oS.WriteText STRING (iIndentIndex, vbTab) & sLine & vbCrLf
Else
oS.WriteText SPACE (nTabSpacing * iIndentIndex) & sLine & vbCrLf
End If
 
If bAdjustIndent Then
iIndentIndex = iIndentIndex + aKey(iIndex)(INDENT_Post_Indent)
If iIndentIndex < 0 Then iIndentIndex = 0
End If
Next
 
oS.Position = 0
BeautifyVBS = oS.ReadText (-1)
oS.Close()
Set oS = Nothing
End Function