'check all arguments is valid path in file system
if wscript.arguments.count <> 0 then
    'create file system object
    set fs_obj = CreateObject("Scripting.FileSystemObject")
    'check path existence and print result
    for i = 0 to wscript.arguments.count - 1
        path = wscript.arguments(i)
        if fs_obj.FolderExists(path) then
            wscript.echo "'" & path & "' - does exist!"
        else
            wscript.echo "'" & path & "' - does not exist!"
        end if
    next
else
    wscript.echo "error: no arguments specified!"
end if
