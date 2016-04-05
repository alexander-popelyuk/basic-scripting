'----------------------------------------------------------
' ls.vbs - list files and directories
'----------------------------------------------------------

'define printing procedure
sub echo(x)
    wscript.echo x
end sub

'define path content printing procedure 
sub print_items(path)
    'get folder object
    set fo = fso.GetFolder(path)
    set files = fo.Files
    set folders = fo.SubFolders

    'print folders
    for each folder in folders
        echo folder.name
    next

    'print files
    for each file in files
        echo file.name
    next
end sub


'create arguments reference
set arg = wscript.arguments
'create file system object
set fso = CreateObject("Scripting.FileSystemObject")

'check arguments count
select case arg.count
    case 0
        'just print content of the current folder
        print_items(".\")
    case 1
        'just print content of the specified location
        print_items(arg(0))
    case else
        'print items for each path argument
        for i = 0 to arg.count - 1
            'print path
            echo vbCRLF & fso.GetAbsolutePathName(arg(i)) & ":"
            'print items for path
            print_items(arg(i))
        next
end select

