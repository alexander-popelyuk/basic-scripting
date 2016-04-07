'----------------------------------------------------------------------
' working with files
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

'get temporary storage path

'create folder
'create file
'write to file
'read from file
'rename file
'copy file
'move file
'delete file
'delete folders
