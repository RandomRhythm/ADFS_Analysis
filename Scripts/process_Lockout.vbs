Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1


Dim DictIP: Set DictIP = CreateObject("Scripting.Dictionary")
Dim dictUserName: Set dictUserName = CreateObject("Scripting.Dictionary")
Dim DictLockSuccess: Set DictLockSuccess = CreateObject("Scripting.Dictionary") 'This account may be compromised. 
Dim dictPwdCount: Set dictPwdCount = CreateObject("Scripting.Dictionary")
Dim dictLastDate: Set dictLastDate = CreateObject("Scripting.Dictionary")
Dim dictInstanceID: Set dictInstanceID = CreateObject("Scripting.Dictionary")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim ArrayEvents
Dim ArrayElines
Dim intANLoc
Dim intIPloc
Dim intPortloc
Dim tmpLogItems
 strFile= "C:\ADFS\512_adfs_02.txt"
 strIDPath = "C:\ADFS\512_adfs_02.log"
 StrOutfile = "C:\ADFS\512_adfs_02.CSV"
 if objFSO.fileexists(strFile) then
  Set objFile = objFSO.OpenTextFile(strFile)

    strEventData = ""
    strData = " "
    if not objFile.AtEndOfStream then 
      'On Error Resume Next
      strData = objFile.ReadAll 

    end if
    if instr(strData, "TimeCreated  : ") > 0 then
      ArrayEvents = split(strData, "TimeCreated  : ")
      msgbox ubound(ArrayEvents) & " events found"
      for each strEvent in ArrayEvents
        if instr(strEvent, vbcrlf) then
          strDate = left(strEvent, 20)
          strUser = getdata(strEvent, vbCr ,"User: " & vbcrlf)
          strIPaddress = getdata(strEvent, vbCr ,"Client IP: " & vbcrlf)
          strPwdCount = getdata(strEvent, vbCr ,"Bad Password Count: " & vbcrlf)
          strLastDate = getdata(strEvent, vbCr ,"nLast Bad Password Attempt: " & vbcrlf)
          
          boolCompromise = False
          if instr(strEvent, "This account may be compromised.") > 0 then 
            boolCompromise = True

          end if
          rowOut = replace( strDate & "," &  strUser & "," & strIPaddress & "," & strPwdCount & "," & strLastDate & "," & boolCompromise ,"               ", "")
          
          
          logdata StrOutfile,rowOut, false
            

         'logdata StrOutfile, "Debug-" & strDate & "," &  strInstanceID & "," & strActivityID & "," & strIP & "," & strUser & "," & strApplication & "," & strUagent & "," & StrSuc & "," & strFail, false

        else  
        
        end if

      next
    else
      msgbox "String not found. Ensure file is in ANSI format"
      msgbox strData
    end if






end if

function LogData(TextFileName, TextToWrite,EchoOn)
Dim strTmpFilName1
Dim strTmpFilName2
TextFileName = RemoveCharsForFname(TextFileName)

Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      fsoLogData.CreateTextFile TextFileName, True
  End If
on error resume next
Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
if err.number <> 0 then
  msgbox "Error writting to " & TextFileName & " perhaps the file is locked?"
  err.number = 0
  Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
  if err.number <> 0 then exit function
end if

on error goto 0
WriteTextFile.WriteLine TextToWrite
WriteTextFile.Close
Set fsoLogData = Nothing
End Function
Function RemoveCharsForFname(TextFileName)
'Remove unsupported characters from file name
strTmpFilName1 = right(TextFileName, len(TextFileName) - instrrev(TextFileName,"\"))
strTmpFilName2 = replace(strTmpFilName1,"/",".")
'TextFileName = replace(TextFileName,"\",".")
strTmpFilName2 = replace(strTmpFilName2,":",".")
strTmpFilName2 = replace(strTmpFilName2,"*",".")
strTmpFilName2 = replace(strTmpFilName2,"?",".")
strTmpFilName2 = replace(strTmpFilName2,chr(34),".")
strTmpFilName2 = replace(strTmpFilName2,"<",".")
strTmpFilName2 = replace(strTmpFilName2,">",".")
strTmpFilName2 = replace(strTmpFilName2,"|",".")
TextFileName = replace(TextFileName,strTmpFilName1,strTmpFilName2)
'will error if file name is to long
If Len(TextFileName) > 255 Then TextFileName = Left(TextFileName, 255)
RemoveCharsForFname = TextFileName
end function

Function GetData(contents, ByVal EndOfStringChar, ByVal MatchString)
MatchStringLength = Len(MatchString)
x= instr(contents, MatchString)

  if X >0 then
    strSubContents = Mid(contents, x + MatchStringLength, len(contents) - MatchStringLength - x +1)
    if instr(strSubContents,EndOfStringChar) > 0 then
      GetData = Mid(contents, x + MatchStringLength, instr(strSubContents,EndOfStringChar) -1)
      'msgbox "success:" & Mid(contents, x + MatchStringLength, instr(Mid(contents, x + MatchStringLength, len(contents) -x),EndOfStringChar) -1)
      exit function
    else
      GetData = Mid(contents, x + MatchStringLength, len(contents) -x -1)
      'msgbox "failed match:" & Mid(contents, x + MatchStringLength, len(contents) -x -1)
      exit function
    end if
    
  end if
GetData = ""

end Function

