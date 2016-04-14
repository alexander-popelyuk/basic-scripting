'---------------------------------------------------------
' List files and folders
'---------------------------------------------------------

' There are two ways to use this program:
'  1. Call without parameters print content of the current
'     directory.
'  2. Call with path parameters - print content of each 
'     path specified.

'create FileSystemObject
set fso = CreateObject("Scripting.FileSystemObject")

'helper procedures
sub print(x)
    WScript.StdOut.Write x
end sub

sub line(x)
    WScript.StdOut.WriteLine x
end sub

sub print_folder_content(path)
    'get folder object
    set fo = fso.GetFolder(path)
    'print subfolders
    for each folder in fo.Subfolders
        line folder.Name
    next
    'print files
    for each file in fo.Files
        line file.Name
    next
end sub

'check argument count
set args = WScript.arguments
if args.count = 0 then
    'print current directory content
    print_folder_content(".")
else
    'print content of each path
    for each arg in args
        if fso.FolderExists(arg) then
            line ""
            line vbCRLF & "'" & fso.GetAbsolutePathName(arg) & "':"
            print_folder_content arg
        else
            line "path does not exist: '" & arg & "'"
        end if
    next
end if

