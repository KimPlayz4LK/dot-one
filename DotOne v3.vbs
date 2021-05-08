on error resume next
Dim fso,result,source,oneFile,oneFileContent,packedFile,packedFileContents,packedFileName,packedFileContent,newPackedFile,unpackedFilePath,toPackContent,currentFolder,currentFolderFiles,currentFolderFile,folderPath,currentSubfolder,selectedFolder
Set fso=CreateObject("Scripting.FileSystemObject")
Sub hexToText(source)
result=""
strSource=Replace(source," ","")
For lp=1 To Len(source) Step 2
result=result+Chr(CByte("&H"&Mid(source,lp,2)))
Next
End Sub

function listSubfolders(selectedFolder)
for each currentSubfolder in selectedFolder.SubFolders
toPackContent="[FILE]"&Replace(currentSubfolder.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[CONTENT]"
oneFile.Write(toPackContent)
for each currentFolderFile in currentSubfolder.Files
source=currentFolderFile.Path
result=""
with CreateObject("ADODB.Stream")
.Type=1
.Open
.LoadFromFile source
for i=0 to Len(fso.OpenTextFile(source,1).ReadAll)-1
.Position=i
result=result&Right(00 & Hex(AscB(.Read(1))),2)
if Err.Number <> 0 Then
Err.Clear
end if
next
end with
toPackContent="[FILE]"&Replace(currentFolderFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[CONTENT]"&result
oneFile.Write(toPackContent)
next
listSubfolders currentSubfolder
next
end function

if WScript.Arguments.Length>0 then
if WScript.Arguments.Length=1 and Right(WScript.Arguments(0),4)=".one" then
Set oneFile=fso.OpenTextFile(WScript.Arguments(0),1,True)
oneFileContent=oneFile.ReadAll
for each packedFile in Split(oneFileContent,"[FILE]")
if inStr(packedFile,"[CONTENT]") then
packedFileContents=Split(packedFile,"[CONTENT]")
packedFileName=fso.GetParentFolderName(WScript.ScriptFullName)&"\"&packedFileContents(0)
if Right(packedFileName,1)="\" then
fso.CreateFolder(packedFileName)
else
set newPackedFile=fso.OpenTextFile(packedFileName,2,true)
hexToText packedFileContents(1)
newPackedFile.Write(result)
newPackedFile.Close
end if
end if
next
if Err.Number <> 0 Then
x=msgbox("Unpacking process done, with errors.",48,"DotOne")
Err.Clear
else
x=msgbox("Unpacking process done.",64,"DotOne")
End If
else
set oneFile=fso.OpenTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\packedFiles.one",2,True)
for each unpackedFilePath in WScript.Arguments
if fso.FolderExists(unpackedFilePath) then
toPackContent="[FILE]"&Replace(unpackedFilePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[CONTENT]"
oneFile.Write(toPackContent)
for each currentFolderFile in fso.GetFolder(unpackedFilePath).Files
source=currentFolderFile.Path
result=""
with CreateObject("ADODB.Stream")
.Type=1
.Open
.LoadFromFile source
for i=0 to Len(fso.OpenTextFile(source,1).ReadAll)-1
.Position=i
result=result&Right(00 & Hex(AscB(.Read(1))),2)
if Err.Number <> 0 Then
Err.Clear
end if
next
end with
toPackContent="[FILE]"&Replace(currentFolderFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[CONTENT]"&result
oneFile.Write(toPackContent)
next
listSubfolders fso.GetFolder(unpackedFilePath)
else
source=unpackedFilePath
result=""
with CreateObject("ADODB.Stream")
.Type=1
.Open
.LoadFromFile source
for i=0 to Len(fso.OpenTextFile(source,1).ReadAll)-1
.Position=i
result=result&Right(00 & Hex(AscB(.Read(1))),2)
if Err.Number <> 0 Then
Err.Clear
end if
next
end with
toPackContent="[FILE]"&Replace(unpackedFilePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[CONTENT]"&result
oneFile.Write(toPackContent)
end if
next
oneFile.Close
if Err.Number <> 0 Then
x=msgbox("Packing process done, with errors.",48,"DotOne")
Err.Clear
else
x=msgbox("Packing process done.",64,"DotOne")
End If
end if
else
x=msgbox("No files detected, try drag'n'dropping files you want to pack or the .ONE file to unpack.",0,"DotOne")
end if