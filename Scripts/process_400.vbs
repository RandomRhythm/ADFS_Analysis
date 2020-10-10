Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1


Dim DictIP: Set DictIP = CreateObject("Scripting.Dictionary")
Dim dictUserName: Set dictUserName = CreateObject("Scripting.Dictionary")
Dim dictClientApp: Set dictClientApp = CreateObject("Scripting.Dictionary")
Dim dictUserAgent: Set dictUserAgent = CreateObject("Scripting.Dictionary")
Dim dictActivityID: Set dictActivityID = CreateObject("Scripting.Dictionary")
Dim dictInstanceID: Set dictInstanceID = CreateObject("Scripting.Dictionary")
Dim dictRelyParty: Set dictRelyParty = CreateObject("Scripting.Dictionary")
Dim dictSucIssue: Set dictSucIssue = CreateObject("Scripting.Dictionary")
Dim dictFailIssue: Set dictFailIssue = CreateObject("Scripting.Dictionary")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim ArrayEvents
Dim ArrayElines
Dim intANLoc
Dim intIPloc
Dim intPortloc
Dim tmpLogItems
 strFile= "C:\ADFS\XXX_adfs_02.txt"
 strIDPath = "C:\ADFS\XXX_adfs_02.log"
 StrOutfile = "C:\ADFS\XXX_adfs_02.CSV"
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
          strInstanceID = getdata(strEvent, vbCr ,"Instance ID: " & vbcrlf)
          if strInstanceID  = "" then strInstanceID  = getdata(strEvent, vbCr ,"Instance ID:  " & vbcrlf)
          strInstanceID = replace(strInstanceID, " ", "")     
          if strInstanceID  = "" then strInstanceID = getdata(strEvent, " " ,"Instance ID: ")
          'if instr(strEvent, "Instance ID: ") then msgbox "strInstanceID=" & strInstanceID
          
          strActivityID = getdata(strEvent, vbCr ,"Activity ID: " & vbcrlf)
          if strActivityID  = "" then strActivityID  = getdata(strEvent, vbCr ,"Activity ID:  " & vbcrlf)
          strActivityID = replace(strActivityID, " ", "")               
   
          if strActivityID  = "" then strActivityID = getdata(strEvent, vbcr ,"Activity ID: ")
          if strActivityID = "00000000-0000-0000-0000-000000000000" then strActivityID = ""
          
          if strActivityID <> "" and dictActivityID.exists(strActivityID) = false then dictActivityID.add strActivityID, strInstanceID

          'set the instance ID so we can track activity
          if strInstanceID  = "" and strActivityID <> "" then strInstanceID = dictActivityID.item(strActivityID)
          'record Activity ID for logging
          if strActivityID <> "" and strInstanceID <> "" and dictInstanceID.exists(strInstanceID) = false then 
            dictInstanceID.add strInstanceID, strActivityID
            'msgbox "adding ID " & strInstanceID & "|" & strActivityID
          end if
          strRely = getdata(strEvent, " " ,"Relying party: ")
          if strRely <> "" and strInstanceID <> "" and dictRelyParty.exists(strInstanceID) = false then 
            dictRelyParty.add strInstanceID, strRely
          elseif strRely <> "" and strInstanceID = "" then
            msgbox "Parsing error 5. InstanceID blank " & strEvent
          end if
          if instr(strEvent, "he Federation Service could not authorize token issuance for caller ") and  dictFailIssue.exists(strInstanceID) = false then dictFailIssue.add strInstanceID, "True"
          if instr(strEvent, "token was successfully issued for the ") and dictSucIssue.exists(strInstanceID) = false then dictSucIssue.add strInstanceID, "True"
          if instr(strEvent, "was successfully authenticated") and dictSucIssue.exists(strInstanceID) = false then dictSucIssue.add strInstanceID, "True"

          
          logdata strIDPath,    strInstanceID & "," & strActivityID, false
          strIP = getdata(strEvent, vbCr ,"http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-forwarded-client-ip " & vbcrlf)
          
          if instr(strEvent, "http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-forwarded-client-ip ") then 
            arrayIPaddresses = split(strEvent, "http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-forwarded")
            for each IpAddress in arrayIPaddresses
              strIP = getdata(IpAddress, vbCr ,"-client-ip " & vbcrlf)
              if instr(strIP, ",") then
                arrayIP = split(strIP, ",")
                for each strTmpIP in arrayIP
                  if replace(strTmpIP, " ", "") <> "172.16.36.42" then
                    strTmpIPD = replace(strTmpIP, " ", "")
                    if strInstanceID <> "" then 
                      if DictIP.exists(strInstanceID) then
                        if instr(DictIP.item(strInstanceID), replace(strTmpIP, " ", "")) = 0 then
                          DictIP.item(strInstanceID) = DictIP.item(strInstanceID) & "|" & replace(strTmpIP, " ", "")
                          'msgbox "Ninja!!" & DictIP.item(strInstanceID) & "|" & replace(strTmpIP, " ", "")
                        end if
                      elseif strTmpIP <> "" then
                        DictIP.add strInstanceID, replace(strTmpIP, " ", "")
                      end if
                    end if
                  end if
                next
              elseif strIP <> "" and  replace(strIP, " ", "") <> "172.16.36.42" and DictIP.exists(strInstanceID) = false then
                DictIP.add strInstanceID, replace(strIP, " ", "")
              elseif strIP <> "" and  replace(strIP, " ", "") <> "172.16.36.42" and DictIP.exists(strInstanceID) = true then
                if instr(DictIP.item(strInstanceID), replace(strIP, " ", "")) = 0  then
                'msgbox "Samurai!!" & DictIP.item(strInstanceID) & "|" & replace(strIP, " ", "")
                'msgbox instr(DictIP.item(strInstanceID), replace(strIP, " ", ""))
                DictIP.item(strInstanceID) = DictIP.item(strInstanceID) & "|" & replace(strIP, " ", "")
                end if
              end if
            next
          end if
          if strIP = "" then strIP = getdata(strEvent, " " ,"X-MS-Forwarded-Client-IP: ")
          if strIP <> "" then 
            if instr(strIP, ",") = 0 then 
              if DictIP.exists(strInstanceID) = false then DictIP.add strInstanceID, strIP
            elseif DictIP.exists(strInstanceID) = false then 
              DictIP.add strInstanceID, left(strIP, instr(strIP, ",") -1)
            
            end if
          end if
          if dictIP.exists(strInstanceID) = false and instr(strEvent, " party '-' was ") then 
            dictIP.add strInstanceID, "Internal"
            'msgbox "internal"
          end if
          
          strApplication = getdata(strEvent, vbCr ,"http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-client-application " & vbcrlf)
          strApplication = replace(strApplication, "               ", "")
          if strApplication = "" then strApplication = getdata(strEvent, " " ,"X-MS-Client-Application: ")
          if strApplication <> "" and dictClientApp.exists(strInstanceID) = false then 
            'msgbox "adding:" & strApplication
            dictClientApp.add strInstanceID, strApplication
          end if

          
          strUagent = getdata(strEvent, vbCr ,"http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-client-user-agent " & vbcrlf)
          strUagent = replace(strUagent, "               ", "")   
          if strUagent = "" then strUagent = getdata(strEvent, " " ,"X-MS-Client-User-Agent: ")
          if strUagent <> "" and dictUserAgent.exists(strInstanceID) = false then dictUserAgent.add strInstanceID, strUagent
          'if instr(strEvent, "http://schemas.microsoft.com/2012/01/requestcontext/claims/x-ms-client-user-agent ") then msgbox "strApplication=" & dictUserAgent.item(strInstanceID)
          
          strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn " & vbcrlf)
          if len(strUser) = 0 and dictUserName.exists(strInstanceID) = false then
            strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/implicitupn " & vbcrlf)
          end if
          if len(strUser) = 0 and dictUserName.exists(strInstanceID) = false then
            strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name " & vbcrlf)
          end if
          if len(strUser) = 0 and dictUserName.exists(strInstanceID) = false then
            strUser = getdata(strEvent, vbCr ,"http://schemas.xmlsoap.org/claims/UPN " & vbcrlf)
          end if
          strUser = replace(strUser, " ", "")      
          'if instr(strEvent, "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn " & vbcr) then msgbox "strUser=" & strUser 
          'msgbox strInstanceID & "|" & strTmpIPD & "|" & strUser
          if strUser <> "" and dictUserName.exists(strInstanceID) = false and strInstanceID <> "" then  dictUserName.add strInstanceID, strUser
          
          if DictIP.exists(strInstanceID) and dictUserName.exists(strInstanceID) then
           if dictIP.item(strInstanceID) = "Internal" or (dictClientApp.exists(strInstanceID) and dictUserAgent.exists(strInstanceID) and (dictSucIssue.exists(strInstanceID) or dictFailIssue.exists(strInstanceID))) then
              StrSuc = "False"
              if dictSucIssue.exists(strInstanceID) then
                StrSuc = dictSucIssue.item(strInstanceID) 
              end if 
              strFail = "False"
              if dictFailIssue.exists(strInstanceID) then 
                strFail = dictFailIssue.item(strInstanceID)
              end if

              logdata StrOutfile, strDate & "," &  strInstanceID & "," & dictInstanceID(strInstanceID) & "," & DictIP.item(strInstanceID) & "," & dictUserName.item(strInstanceID) & "," & dictClientApp.item(strInstanceID) & "," & dictUserAgent.item(strInstanceID) & "," & StrSuc & "," & strFail, false
            end if
          end if
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

