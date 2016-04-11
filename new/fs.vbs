'----------------------------------------------------------------------
' working with file system
'----------------------------------------------------------------------

'all file manipulations start with creating the 'FileSystemObject'
set fso = CreateObject("Scripting.FileSystemObject")

'create sub for debug output
sub print(x)
    WScript.StdOut.Write x
end sub

'create function to convert drive type to readable string
function drvType2str(drvType)
    select case drvType
        Case 0: drvType2str = "Unknown"
        Case 1: drvType2str = "Removable"
        Case 2: drvType2str = "Fixed"
        Case 3: drvType2str = "Network"
        Case 4: drvType2str = "CD-ROM"
        Case 5: drvType2str = "RAM Disk"
    end select
end function

'list mounted drives
for each drv in fso.Drives
    type_str = drvType2str(drv.DriveType)
    print drv & " - "
    if type_str = "Fixed" then
        print(drv.VolumeName)
    else
        print type_str
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

print "Windows Folder is: '" & win & "'" & vbCRLF
print "System Folder is: '" & sys & "'" & vbCRLF
print "Temporary Folder is: '" & tmp & "'" & vbCRLF

'create folder
fname = "vbs_folder"
fpath = fso.BuildPath(tmp, fname)
print "folder path is: " & fpath & vbCRLF
set fobj = fso.CreateFolder(fpath)
    ' stop until describe error handling
if IsNull(fobj) then
    print "unable to create folder!"
else
    print "folder successfully created!"
end if

'create file
'write to file
'read from file
'rename file
'copy file
'move file
'delete file
'delete folders
