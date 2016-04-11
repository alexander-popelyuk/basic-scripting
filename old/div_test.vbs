for i = 0 to wscript.arguments.count - 1
    arg = wscript.arguments(i)
    msg = ""
    if (arg mod 2) = 0 then
        msg = msg & "even" & vbCRLF
    else
        msg = msg & "odd" & vbCRLF
    end if
    MsgBox msg, vbOkOnly, arg & " is:"
next
