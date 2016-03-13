set fso = CreateObject("Scripting.FileSystemObject")
if wscript.arguments.count then
    folder = wscript.arguments(0)
else
    folder = "C:\Temp"
end if
set FolderObject = fso.GetFolder(folder).Files
FileList = ""
for each file in FolderObject
    FileList = FileList & """" & file.name & """" & ";" & vbCRLF
next
MsgBox "The files in " & """" & folder & """" & " are:" & vbCRLF & FileList

