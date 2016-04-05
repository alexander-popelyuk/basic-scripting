sub msg(text)
    MsgBox text, vbOkOnly, "info"
end sub

sub echo(text)
    wscript.echo text
end sub

sub line(text)
    wscript.echo text & vbCRLF
end sub

Function cit(text)
    cit = "'" & text & "'"
end Function

sub eol()
    wscript.echo vbCRLF
end sub

str1 = "this is first string"

pos1 = InStr(str1, "is")

pos2 = InStrRev(str1, "is")
echo vbCRLF & "original string: '" & str1 & "'" & vbCRLF
echo "pos1 = " & pos1 & vbCRLF & "pos2 = " & pos2 & vbCRLF _
    & "left 4 = '" & Left(str1, 4) &"'" & vbCRLF _
    & "middle 6, 8 = '" & Mid(str1, 6, 8) &"'" & vbCRLF _
    & "right 6 = '" & Right(str1, 6) &"'"

str2 = "    here a text        "


echo vbCRLF & "original string: '" & str2 & "'" & vbCRLF

echo "Trim = " & cit(Trim(str2))
echo "Ltrim = " & cit(Ltrim(str2))
echo "Rtrim = " & cit(Rtrim(str2))
echo "Ucase = " & cit(Ucase(str2))
echo "Lcase = " & cit(Lcase(str2))
echo "Len = " & Len(str2)
