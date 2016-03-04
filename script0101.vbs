'script0101.vbs
'total up space used in given directory

dir = "C:\"

set Fsys = CreateObject("Scripting.FileSystemObject")
totsize = 0
for each file in Fsys.GetFolder(dir).Files
    totsize = totsize + file.size
next
wscript.echo "The total size of the files in", dir, "is", totsize, "bytes"

