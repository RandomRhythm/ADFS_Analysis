Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1


Dim DictIP: Set DictIP = CreateObject("Scripting.Dictionary")
Dim dictUserName: Set dictUserName = CreateObject("Scripting.Dictionary")
Dim DictLocked: Set DictLocked = CreateObject("Scripting.Dictionary") 'This account may be compromised. 
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
 strFile= "C:\ADFS\5254_adfs_01.txt"
 strIDPath = "C:\ADFS\5254_adfs_01_all.csv"
 StrOutfile = "C:\ADFS\5254_adfs_01.CSV"
 if objFSO.fileexists(strFile) then
  Set objFile = objFSO.OpenTextFile(strFile)

    strEventData = ""
    strData = " "
    if not objFile.AtEndOfStream then 
      'On Error Resume Next
      strData = objFile.ReadAll 

    end if
    
    boolBadPass = False
    boolLocked = False
    boolSucess = False
    boolSTSLock = False
    boolExpired = False
    boolDisabled = False
    boolCheckNext = False
    if instr(strData, "TimeCreated  : ") > 0 then
      ArrayEvents = split(strData, "TimeCreated  : ")
      msgbox ubound(ArrayEvents) & " events found"
      strPreviousEvent = ""
      for each strEvent in ArrayEvents
        if instr(strEvent, vbcrlf)  > 0 and strPreviousEvent <> strEvent then
          if boolCheckNext = False then           
          strDate = getdata(strEvent, vbcr ,"Currenttime: ")
          strUser = getdata(strEvent, " " ,"User: ")
          strPwdCount = getdata(strEvent, " " ,"BadPwdCount: ")
          strLastDate = getdata(strEvent, "O" ,"LastBadPasswordAttempt: ")
          strLastDate = replace(strLastDate, vbcrlf, "")
          end if
          if instr(strEvent, "The user name or password is incorrect") > 0 then 
            boolBadPass = True
            strBadPwdUser = rgetdata(strEvent, vblf, "-The user name or password is incorre")
            if strBadPwdUser = "" then
              strBadPwdUser = getdata(strEvent, " ", "System.IdentityModel.Tokens.SecurityTokenValidationException: ")
            end if
            strBadPwdUser = replace(strBadPwdUser, " ", "")
            
          end if
          if instr(strEvent, "is currently locked out") > 0 then 
            boolLocked = True
          end if
          if instr(strEvent, "exceeded the STS max bad password count") > 0 then boolSTSLock = True
          if boolSTSLock = True or boolLocked = True then
            if strLockUser = "" then strLockUser = rgetdata(strEvent, vblf, "-The referenced account")
            if strLockUser = "" then
              strLockUser = getdata(strEvent, " ", "System.IdentityModel.Tokens.SecurityTokenValidationException: ")
            end if
            strLockUser = replace(strLockUser, " ", "")
            if DictLocked.exists( strLockUser) = false then DictLocked.add strLockUser, True
          end if
          if instr(strEvent, "This user can't sign in because this account is currently disable") > 0 then boolDisabled = True
            if boolDisabled = True or boolLocked = True then
            if strDisableUser = "" then strDisableUser = rgetdata(strEvent, vblf, "--This user can't sign in because this account is currently disable")
            if strDisableUser = "" then
              strDisableUser = getdata(strEvent, " ", "System.IdentityModel.Tokens.SecurityTokenValidationException: ")
            end if
            strDisableUser = replace(strDisableUser, " ", "")
          end if
          if instr(strEvent, "password for this account has expired") > 0 then 
            boolExpired = True
            if strExpireUser = "" Then strExpireUser = rgetdata(strEvent, vblf, "-The password for this account has expire")
            if strExpireUser = "" then
              strExpireUser = getdata(strEvent, " ", "System.IdentityModel.Tokens.SecurityTokenValidationException: ")
            end if
            strExpireUser = replace(strExpireUser, " ", "")
          end if
          if instr(strEvent, "LogSuccessAuthenticationInfo:") > 0 and _
           instr(strEvent, "identitymodel/tokens/UserName") > 0 then boolSucess = True
          if strDate <> "" and strPwdCount <> "" then
            if instr(strEvent, strLockUser) = False and DictLocked.exists(strUser) = False and boolLocked = True then 
              boolLocked = False
              'msgbox "User " & strLockUser & " doesn't match event " & strEvent
              logdata strIDPath & ".log",strLockUser & " <> " & strEvent, false
            end if
            if instr(strEvent, strBadPwdUser) = False and boolBadPass = True then 
              boolBadPass = False
            end if
            if instr(strEvent, strExpireUser) = False and boolExpired = True then 
              boolExpired = False
            end if
            if instr(strEvent, strLockUser) = True and boolSTSLock = True then 
              boolSucess = False
            end if
            if boolCheckNext = True or boolLocked = True or boolBadPass = True or boolExpired = True or boolSTSLock = True or boolSucess = True then            
              if boolLocked = True or boolBadPass = True or boolExpired = True or boolSTSLock = True then boolSucess = False
              'Date Time	User	Bad Password Count	Last Bad Password	Bad Password	STS Lock	Locked	Success Auth	Expired Password	Disabled


              rowOut = replace( strDate & "," &  strUser & "," & strPwdCount & "," & strLastDate & "," & boolBadPass & "," & boolSTSLock & "," & boolLocked & "," & boolSucess & "," & boolExpired & "," & boolDisabled,"                ", "")
              if dictUserName.exists(strUser) then
                if int(dictUserName.item(strUser)) < int(strPwdCount) then
                  logdata StrOutfile,rowOut & "," & dictUserName.item(strUser), false
                end if
              elseif isnumeric(strPwdCount) then
                dictUserName.add strUser, strPwdCount
              end if
              'if boolLocked = True or boolBadPass = True or boolExpired = True or boolSTSLock = True or boolSucess = True then            
                logdata strIDPath,rowOut, false
              'end if
              boolBadPass = False
              boolLocked = False
              boolSucess = False
              boolSTSLock = False
              boolExpired = False
              boolDisabled = False
              strLockUser = ""
              strExpireUser = ""
              boolCheckNext = False
              DictLocked.RemoveAll
            else
              boolCheckNext = True
            end if
          end if
          strPreviousEvent = strEvent
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

Function rGetData(contents, ByVal EndOfStringChar, ByVal MatchString)
MatchStringLength = Len(MatchString)
x= instrRev(contents, MatchString)

  if X >0 then
    if instrRev(left(contents, x),EndOfStringChar) > 0 then
      rGetData = Mid(contents, instrRev(left(contents, x),EndOfStringChar) +len(EndOfStringChar),x - instrRev(left(contents, x),EndOfStringChar) -len(EndOfStringChar))
      'msgbox "success:" & Mid(contents, instrRev(left(contents, x),EndOfStringChar) +len(EndOfStringChar),x -instrRev(left(contents, x),EndOfStringChar)-len(EndOfStringChar) )
      exit function
    else
      rGetData = left(contents,x)
      'msgbox "failed match:" & left(contents,x -1)
      exit function
    end if
    
  end if
rGetData = ""

end Function
