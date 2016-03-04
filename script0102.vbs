'script0102.vbs
'total up space used in directory named on the command line

dir = wscript.arguments(0)

set fsys = createObject("Scripting.FileSystemObject")
totsize = 0
for each file in fsys.GetFolder(dir).Files
    totsize = totsize + file.size
next
wscript.echo "The total size of the files in", dir, "is", totsize

