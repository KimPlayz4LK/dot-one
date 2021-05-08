on error resume next
ENCODE=true
Set fso=CreateObject("Scripting.FileSystemObject")
Function Base64Encode(inData)
Const Base64="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim cOut,sOut,I
For I=1 To Len(inData) Step 3
Dim nGroup,pOut,sGroup
nGroup=&H10000 * Asc(Mid(inData,I,1)) + _
&H100 * MyASC(Mid(inData,I + 1,1)) + MyASC(Mid(inData,I + 2,1))
nGroup=Oct(nGroup)
nGroup=String(8 - Len(nGroup),"0") & nGroup
pOut=Mid(Base64,CLng("&o" & Mid(nGroup,1,2)) + 1,1) + _
Mid(Base64,CLng("&o" & Mid(nGroup,3,2)) + 1,1) + _
Mid(Base64,CLng("&o" & Mid(nGroup,5,2)) + 1,1) + _
Mid(Base64,CLng("&o" & Mid(nGroup,7,2)) + 1,1)
sOut=sOut + pOut
Next
Select Case Len(inData) Mod 3
Case 1: '8 bit final
sOut=Left(sOut,Len(sOut) - 2) + "=="
Case 2: '16 bit final
sOut=Left(sOut,Len(sOut) - 1) + "="
End Select
Base64Encode=sOut
End Function
Function MyASC(OneChar)
If OneChar="" Then MyASC=0 Else MyASC=Asc(OneChar)
End Function
Function Base64Decode(ByVal base64String)
Const Base64="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim dataLength,sOut,groupBegin
base64String=Replace(base64String,vbCrLf,"")
base64String=Replace(base64String,vbTab,"")
base64String=Replace(base64String," ","")
dataLength=Len(base64String)
If dataLength Mod 4 <> 0 Then
Err.Raise 1,"Base64Decode","Bad Base64 string."
Exit Function
End If
For groupBegin=1 To dataLength Step 4
Dim numDataBytes,CharCounter,thisChar,thisData,nGroup,pOut
numDataBytes=3
nGroup=0
For CharCounter=0 To 3
thisChar=Mid(base64String,groupBegin + CharCounter,1)
If thisChar="=" Then
numDataBytes=numDataBytes - 1
thisData=0
Else
thisData=InStr(1,Base64,thisChar,vbBinaryCompare) - 1
End If
If thisData=-1 Then
Err.Raise 2,"Base64Decode","Bad character In Base64 string."
Exit Function
End If
nGroup=64 * nGroup + thisData
Next
nGroup=Hex(nGroup)
nGroup=String(6 - Len(nGroup),"0") & nGroup
pOut=Chr(CByte("&H" & Mid(nGroup,1,2))) + _
Chr(CByte("&H" & Mid(nGroup,3,2))) + _
Chr(CByte("&H" & Mid(nGroup,5,2)))
sOut=sOut & Left(pOut,numDataBytes)
Next
Base64Decode=sOut
End Function

Function subfolder(folder)
For Each Subfolder in folder.SubFolders
content="[START_OF_NEW_OBJECT]"&Replace(Subfolder.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[START_OF_OBJECT_CONTENT]"
oneFile.Write(content)
Set objFolder=fso.GetFolder(Subfolder.Path)
Set colFiles=objFolder.Files
For Each objFile in colFiles
set file=fso.OpenTextFile(objFile.Path,1)
fileContent=file.ReadAll()
file.Close()
if ENCODE=true then
content="[START_OF_NEW_OBJECT]"&Replace(objFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&Base64Encode(fileContent)
else
content="[START_OF_NEW_OBJECT]"&Replace(objFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&fileContent
end if
oneFile.Write(content)
Next
subfolder(Subfolder)
Next
End Function

if WScript.Arguments.Length>0 then
if Right(WScript.Arguments(0),4)=".one" then
Set oneFile=fso.OpenTextFile(WScript.Arguments(0),1)
fileContent=oneFile.ReadAll
if inStr(Split(fileContent,"[START_OF_NEW_OBJECT]")(0),"True") then
ENCODE=true
else
ENCODE=false
end if
for each fileStuff in Split(fileContent,"[START_OF_NEW_OBJECT]")
if inStr(fileStuff,"[START_OF_OBJECT_CONTENT]") then
fileStuff=Split(fileStuff,"[START_OF_OBJECT_CONTENT]")
if Right(fileStuff(0),1)="\" then
if not fso.FolderExists(fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0)) then
fso.CreateFolder(fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0))
if Err.Number <> 0 Then
x=msgbox("An error occured while trying to create the folder"&vbCrLf&fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0)&vbCrLf&vbCrLf&"Error: "&Err.Description,16,"DotOne")
Err.Clear
End If
end if
else
set file=fso.CreateTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0))
if Err.Number <> 0 Then
x=msgbox("An error occured while trying to unpack the file"&vbCrLf&fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0)&vbCrLf&vbCrLf&"Error: "&Err.Description,16,"DotOne")
Err.Clear
End If
if ENCODE=true then
toWrite=Base64Decode(fileStuff(1))
else
toWrite=fileStuff(1)
end if
'toWrite=Left(toWrite,Len(toWrite)-1)
file.Write(toWrite)
file.Close()
end if
end if
next
x=msgbox("Unpacking process done.",64,"DotOne")
else
set oneFile=fso.CreateTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\packedFiles.one")
if Err.Number <> 0 Then
x=msgbox("An error occured while trying to pack your files"&vbCrLf&vbCrLf&"Error: "&Err.Description,16,"DotOne")
Err.Clear
End If
x=msgbox("Would you like to encode your files? (recommended if you have plain-text files, but your files could be corrupted if not plain-text)",4,"DotOne")
if x=vbNo then
ENCODE=false
end if
oneFile.Write(ENCODE)
for each filePath in WScript.Arguments
if fso.FolderExists(filePath) then
content="[START_OF_NEW_OBJECT]"&Replace(filePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[START_OF_OBJECT_CONTENT]"
oneFile.Write(content)

Set folder=fso.GetFolder(filePath)
Set files=folder.Files

For Each subFile in files
set file=fso.OpenTextFile(subFile.Path,1)
fileContent=file.ReadAll()
file.Close()
if ENCODE=true then
content="[START_OF_NEW_OBJECT]"&Replace(subFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&Base64Encode(fileContent)
else
content="[START_OF_NEW_OBJECT]"&Replace(subFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&fileContent
end if
oneFile.Write(content)
Next
subfolder(folder)
else
Set file=fso.OpenTextFile(filePath,1)
if Err.Number <> 0 Then
x=msgbox("An error occured while trying to pack the file"&vbCrLf&filePath&vbCrLf&vbCrLf&"Error: "&Err.Description,16,"DotOne")
Err.Clear
End If
fileContent=file.ReadAll
if ENCODE=true then
content="[START_OF_NEW_OBJECT]"&Replace(filePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&Base64Encode(fileContent)
else
content="[START_OF_NEW_OBJECT]"&Replace(filePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&fileContent
end if
oneFile.Write(content)
end if
next
oneFile.Close()
x=msgbox("Packing process done.",64,"DotOne")
end if
else
x=msgbox("Please drag'n'drop the .ONE file to unpack or files to pack.")
end if