Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1


Dim DictIP: Set DictIP = CreateObject("Scripting.Dictionary")
Dim dictUserName: Set dictUserName = CreateObject("Scripting.Dictionary")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim ArrayEvents
Dim ArrayElines
Dim intANLoc
Dim intIPloc
Dim intPortloc
Dim tmpLogItems
 strFile= "C:\ADFS\logonsoff.txt"

 StrOutfile = "C:\ADFS\2_event.txt"
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
    
      for each strEvent in ArrayEvents
        if instr(strEvent, vbcrlf) then
          strInstanceID = getdata(strEvent, vbCr ,"Instance ID: " & vbcrlf)
          strInstanceID = replace(strInstanceID, " ", "")        
          strIP = getdata(strEvent, vbCr ,"http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-forwarded-client-ip " & vbcrlf)
          'if instr(strEvent, "http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-forwarded-client-ip ") then msgbox "IP=" & strIP 
          if instr(strIP, ",") then
            arrayIP = split(strIP, ",")
            for each strTmpIP in arrayIP
              if replace(strTmpIP, " ", "") <> "172.16.36.42" then
                strTmpIPD = replace(strTmpIP, " ", "")
                if strInstanceID <> "" then 
                  if DictIP.exists(strInstanceID) then
                    if instr(DictIP.item(strInstanceID), replace(strTmpIP, " ", "")) = 0 then
                      DictIP.item(strInstanceID) = DictIP.item(strInstanceID) & "|" & replace(strTmpIP, " ", "")
                    end if
                  elseif strTmpIP <> "" then
                    DictIP.add strInstanceID, replace(strTmpIP, " ", "")
                  end if
                end if
              end if
            next
          elseif strIP <> "" and  strIP <> "172.16.36.42" and DictIP.exists(strInstanceID) = false then
            DictIP.add strInstanceID, replace(strIP, " ", "")
          elseif strIP <> "" and  strIP <> "172.16.36.42" and instr(DictIP.item(strInstanceID), replace(strIP, " ", "")) = 0 then
            DictIP.item(strInstanceID) = DictIP.item(strInstanceID) & "|" & replace(strIP, " ", "")
          end if
          
          strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn " & vbcrlf)
          if len(strUser) = 0 and dictUserName.exists(strInstanceID) = false then
            strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/implicitupn " & vbcrlf)
          end if
          if len(strUser) = 0 and dictUserName.exists(strInstanceID) = false then
            strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name " & vbcrlf)
          end if
          strUser = replace(strUser, " ", "")      
          'if instr(strEvent, "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn " & vbcr) then msgbox "strUser=" & strUser 
          'msgbox strInstanceID & "|" & strTmpIPD & "|" & strUser
          if strUser <> "" and dictUserName.exists(strInstanceID) = false and strInstanceID <> "" then  dictUserName.add strInstanceID, strUser
          
          if DictIP.exists(strInstanceID) and dictUserName.exists(strInstanceID) then
          
            logdata StrOutfile, strInstanceID & "," & DictIP.item(strInstanceID) & "," & dictUserName.item(strInstanceID), false
          end if
         

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

