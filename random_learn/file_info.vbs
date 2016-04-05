'consider all arguments as path to some file
'and get information about all of them

set arg = wscript.arguments

if arg.count <> 0 then
    set fso = CreateObject("Scripting.FileSystemObject")
    for i = 0 to arg.count - 1
        path = arg(i)
        if fso.FileExists(path) then
            set file = fso.GetFile(path)
            wscript.echo path & " - file exist, size: " & _
                file.size & " byte(s), created: " & file.DateCreated
        else
            wscript.echo path & " - file not found!"
        end if
    next
else
    wscript.echo "error: no arguments specified!"
end if
