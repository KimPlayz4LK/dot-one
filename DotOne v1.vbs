on error resume next
Set fso=CreateObject("Scripting.FileSystemObject")
if WScript.Arguments.Length>0 then
if Right(WScript.Arguments(0),4)=".one" then
Set oneFile=fso.OpenTextFile(WScript.Arguments(0),1)
fileContent=oneFile.ReadAll
for each fileStuff in Split(fileContent,"[START_OF_NEW_OBJECT]")
if inStr(fileStuff,"[START_OF_OBJECT_CONTENT]") then
fileStuff=Split(fileStuff,"[START_OF_OBJECT_CONTENT]")
if Right(fileStuff(0),1)="\" then
if not fso.FolderExists(fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0)) then
fso.CreateFolder(fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0))
end if


else
set file=fso.CreateTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\"&fileStuff(0))
file.WriteLine(fileStuff(1))
file.Close()
end if
end if
next
else
set oneFile=fso.CreateTextFile(fso.GetParentFolderName(WScript.ScriptFullName)&"\packedFiles.one")
for each filePath in WScript.Arguments
if fso.FolderExists(filePath) then
content="[START_OF_NEW_OBJECT]"&Replace(filePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"\[START_OF_OBJECT_CONTENT]"
oneFile.WriteLine(content)

Set folder=fso.GetFolder(filePath)
Set files=folder.Files
For Each subFile in files
set file=fso.OpenTextFile(subFile.Path,1)
fileContent=file.ReadAll()
file.Close()
content="[START_OF_NEW_OBJECT]"&Replace(subFile.Path,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&fileContent
oneFile.WriteLine(content)
Next

else
Set file=fso.OpenTextFile(filePath,1)
fileContent=file.ReadAll
content="[START_OF_NEW_OBJECT]"&Replace(filePath,fso.GetParentFolderName(WScript.ScriptFullName)&"\","")&"[START_OF_OBJECT_CONTENT]"&fileContent
oneFile.WriteLine(content)
end if
next
oneFile.Close()
end if
else
x=msgbox("Please drag'n'drop the .ONE file to unpack or files to pack.")
end if