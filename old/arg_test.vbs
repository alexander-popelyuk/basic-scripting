REM this script test it's arguments

msgbox "got " & wscript.arguments.count & " arguements!", vbOkOnly, "Info"

for i = 0 to wscript.arguments.count - 1
    dim msg
    select case wscript.arguments(i)
        case 23:
            msg = "it's 23"
        case 2:
            msg = "it's 2"
        case 3:
            msg = "it's 3"
        case 5, 7:
            msg = "5 or 7"
        case else:
            msg = "other number"
    end select
    msgbox msg, vbOkOnly, "Argument " & i + 1 & ":"
next

