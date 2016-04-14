'----------------------------------------------------------------------
' working with file system
'----------------------------------------------------------------------

'all file manipulations start with creating the 'FileSystemObject'
set fso = CreateObject("Scripting.FileSystemObject")

'create sub for debug output
sub print(x)
    WScript.StdOut.Write x
end sub

sub line(x)
    WScript.StdOut.WriteLine x
end sub

'drive types codes
const dtUnknown   = 0
const dtRemovable = 1
const dtFixed     = 2
const dtNetwork   = 3
const dtCdRom     = 4
const dtRamDisk   = 5

'create function to convert drive type to readable string
function drvType2str(drvType)
    select case drvType
        Case dtUnknown   : drvType2str = "Unknown"
        Case dtRemovable : drvType2str = "Removable"
        Case dtFixed     : drvType2str = "Fixed"
        Case dtNetwork   : drvType2str = "Network"
        Case dtCdRom     : drvType2str = "CD-ROM"
        Case dtRamDisk   : drvType2str = "RAM Disk"
    end select
end function

'list mounted drives
line "Available drives is:"
for each drv in fso.Drives
    print drv & " - "
    if drv.DriveType = dtFixed then
        print(drv.VolumeName)
    else
        print drvType2str(drv.DriveType)
    end if
    print vbCRLF
next

'get special folders paths
const sfWindowsFolder   = 0
const sfSystemFolder    = 1
const sfTemporaryFolder = 2

win = fso.GetSpecialFolder(sfWindowsFolder)
sys = fso.GetSpecialFolder(sfSystemFolder)
tmp = fso.GetSpecialFolder(sfTemporaryFolder)

line "Windows Folder is: '" & win & "'"
line "System Folder is: '" & sys & "'"
line "Temporary Folder is: '" & tmp & "'"

'create folder
fname = "test_folder"
fpath = fso.BuildPath(fso.GetAbsolutePathName("."), fname)
line "Try to create folder: " & fpath
on error resume next
set fobj = fso.CreateFolder(fpath)
if Err.Number <> 0 then
    if fso.FolderExists(fpath) then
        set fobj = fso.GetFolder(fpath)
        if IsObject(fobj) then Err.Clear() end if
    end if
end if
if Err.Number <> 0 then
    line "unable to create folder"
    line Err.Description
    WScript.Quit 1
else
    line "folder successfully created!"
end if
on error goto 0

'create temporary file
tpath = fso.BuildPath(fpath, fso.GetTempName())
set tfile = fso.CreateTextFile(tpath, True)

'write to file
line "write to '" & tpath & "'..."
tfile.WriteLine "first test line"
tfile.WriteLine "second test line"
tfile.Close

'read from file
const ForReading = 1, ForAppending = 8
set tfile = fso.OpenTextFile(tpath, ForReading)
line "read back from '" & tpath & "':"
do
    line tfile.ReadLine 
loop until tfile.AtEndOfStream
tfile.Close

'rename file
const new_name = "test_file.txt"
set tfile = fso.GetFile(tpath)
line "rename '" & tfile.Name & "' to '" & new_name & "'"
tfile.Name = new_name

'copy file
const copy_name = "test_file_copy.txt"
copy_path = fso.BuildPath(fpath, copy_name)
line "copy file '" & tfile.Name & "' to '" & copy_name & "'"
tfile.Copy(copy_path)

'move files
const subfolder_name = "test_subfolder"
move_path = fso.BuildPath(fpath, subfolder_name & "\")
line "move '" & copy_path & "' to '" & move_path & "'"
fso.CreateFolder(move_path)
fso.MoveFile copy_path, move_path

'delete file
tpath = fso.BuildPath(fpath, new_name)
line "delete file '" & tpath & "'"
fso.DeleteFile tpath

'delete folders
del_path = fso.BuildPath(fpath, subfolder_name)
line "delete folder '" & del_path & "'"
fso.DeleteFolder del_path
line "delete folder '" & fpath & "'"
fso.DeleteFolder fpath

