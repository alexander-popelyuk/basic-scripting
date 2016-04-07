'----------------------------------------------------------------------
' user input/output
'----------------------------------------------------------------------

'object.Echo [Arg1] [,Arg2] [,Arg3] ... 
WScript.Echo "single string argument"
'force space between arguments
WScript.Echo "first argument", "second argument"
'left arguments as they are
WScript.Echo "chain multi" & "ple strings together"

'MsgBox(prompt[, buttons][, title][, helpfile, context])
MsgBox "Message text", vbOkOnly, "Message title"
'use result of function working
r = MsgBox("Accept or decline", vbOkCancel, "Please, choose!")
if r = vbOk then
    rs = "accepted"
else
    rs = "declined"
end if
WScript.Echo "user was " & rs & " your offer!"

'InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])
v = InputBox("Type a value", "Input Box", "0")
MsgBox "user was entered '" & v & "' into the input box"

'wscript.stdin, wscript.stdout, wscript.stderr (cscript.exe only)
WScript.StdOut.WriteLine "input something here and press enter:"
line = WScript.StdIn.ReadLine()
WScript.StdOut.Write "'" & line
WScript.StdOut.Write "' was read!" & vbCRLF
WScript.StdErr.WriteLine "write something to error stream..."
'per character read from StdIn
do
    ch = WScript.StdIn.Read(1)
    WScript.StdOut.Write ch
    if ch = vbLF and old_ch = vbCR then exit do
    old_ch = ch
loop

'get arguments object
set args = WScript.arguments
'print argument count
if args.count = 0 then
    WScript.StdOut.WriteLine "no arguments specified!"
else
    WScript.StdOut.WriteLine "receive " & args.count & " argument(s)"
end if
'print script arguments
for i = 0 to args.count - 1
    WScript.StdOut.WriteLine "arg " & i & " = '" & args(i) & "'"
next

