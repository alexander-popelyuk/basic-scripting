'ask user for answer
answer = MsgBox("confirm or refuse", _
        vbOkCancel + vbCritical, _
        "please, choose")

'print received answer
select case answer
    case vbOk:
        wscript.echo "'Ok' was pressed"
    case vbCancel:
        wscript.echo "'Cancel' was pressed"
    case else:
        wscript.echo "something else was pressed"
end select

