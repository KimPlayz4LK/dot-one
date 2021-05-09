on error resume next
Dim fso,result,source,oneFile,oneFileContent,packedFile,packedFileContents,packedFileName,packedFileContent,newPackedFile,unpackedFilePath,toPackContent,currentFolder,currentFolderFiles,currentFolderFile,folderPath,currentSubfolder,selectedFolder,pkgName,temp,tempfile,temp2
Set fso=CreateObject("Scripting.FileSystemObject")
Set shell=CreateObject("WScript.Shell")
Sub hexToText(source)
result=""
strSource=Replace(source," ","")
For lp=1 To Len(source) Step 2
result=result+Chr(CByte("&H"&Mid(source,lp,2)))
Next
End Sub
function listSubfolders(selectedFolder)
for each currentSubfolder in selectedFolder.SubFolders
toPackContent="[F]"&Replace(currentSubfolder.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[C]"
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
toPackContent="[F]"&Replace(currentFolderFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[C]"&result
oneFile.Write(toPackContent)
next
listSubfolders currentSubfolder
next
end function
if WScript.Arguments.Length>0 then
if WScript.Arguments.Length=1 then
set tempfile=fso.OpenTextFile(WScript.Arguments(0))
starting=tempfile.ReadAll
if Err.Number <> 0 Then
Err.Clear
end if
else
starting=""
end if
goodStarting=InStr(1,starting,"DotOne package")=1
if WScript.Arguments.Length=1 and goodStarting then
Set oneFile=fso.OpenTextFile(WScript.Arguments(0),1,True)
oneFileContent=oneFile.ReadAll
temp=Split(oneFileContent,"[F]")
temp2=Split(temp(0),vbCrLf)
pkgName=temp2(1)
if Err.Number <> 0 Then
pkgName="packed-files"
Err.Clear
end if
folderName=fso.GetParentFolderName(WScript.ScriptFullName)&"\"&pkgName
fso.CreateFolder(folderName)
if Err.Number <> 0 Then
folderName=fso.GetParentFolderName(WScript.ScriptFullName)&"\"&"unpacked-files"
fso.CreateFolder(folderName)
Err.Clear
end if
for each packedFile in Split(oneFileContent,"[F]")
if inStr(packedFile,"[C]") then
packedFileContents=Split(packedFile,"[C]")
packedFileName=folderName&"\"&packedFileContents(0)
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
x=msgbox("Unpacking process done, with errors."&vbCrLf&"Package name: "&pkgName,48+4096,"DotOne")
Err.Clear
else
x=msgbox("Unpacking process done."&vbCrLf&"Package name: "&pkgName,64+4096,"DotOne")
End If
shell.run fso.GetAbsolutePathName(folderName)
else
pkgName=inputbox("Enter a package name","DotOne v4","packedFiles")
pkgName=replace(pkgName," ","-")
pkgName=replace(pkgName,"\","")
pkgName=replace(pkgName,"/","")
pkgName=replace(pkgName,":","")
pkgName=replace(pkgName,"*","")
pkgName=replace(pkgName,"?","")
pkgName=replace(pkgName,"""","")
pkgName=replace(pkgName,"<","")
pkgName=replace(pkgName,">","")
pkgName=replace(pkgName,"|","")
if len(pkgName)>50 then
pkgName=left(pkgName,50)
end if
set oneFile=fso.OpenTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\packed-"&pkgName&".txt",2,True)
if Err.Number <> 0 Then
set oneFile=fso.OpenTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\packed-files.txt",2,True)
Err.Clear
end if
oneFile.Write("DotOne package"&vbCrLf)
oneFile.Write(pkgName&vbCrLf)
for each unpackedFilePath in WScript.Arguments
if fso.FolderExists(unpackedFilePath) then
toPackContent="[F]"&Replace(unpackedFilePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[C]"
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
toPackContent="[F]"&Replace(currentFolderFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[C]"&result
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
toPackContent="[F]"&Replace(unpackedFilePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[C]"&result
oneFile.Write(toPackContent)
end if
next
oneFile.Close
if Err.Number <> 0 Then
x=msgbox("Packing process done, with errors."&vbCrLf&"Package name: "&pkgName,48+4096,"DotOne")
Err.Clear
else
x=msgbox("Packing process done."&vbCrLf&"Package name: "&pkgName,64+4096,"DotOne")
End If
end if
else
x=msgbox("No files detected, try drag'n'dropping files you want to pack or the DotOne package to unpack.",16+4096,"DotOne")
end if