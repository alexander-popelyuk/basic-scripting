'   User interface
'   ===================================================================

'   This file describe real-time user interaction. Here described 5
'   ways of interaction:
'       1.  using `Echo()` method of the `WScript` object;
'       2.  using `MsgBox()` global function;
'       3.  using `InputBox()` global function;
'       4.  using standard input, output and error streams;
'       5.  using arguments passed to script through command line
'           invoking.

'   1.  Echo method
'   ----------------------------------------------------------------------

'   WScript.Echo [arg1] [,arg2] [,arg3] ... 

'   Print text to standard output or create new message box for each
'   call depending of interpreter used: _cscript.exe_ or _wscript.exe_.

'print single argument
WScript.Echo "single string argument"
'force space between arguments
WScript.Echo "first argument", "second argument"
'to eliminate spaces, chain multiple string
WScript.Echo "chain multi" & "ple strings together"

'   2.  Message Box
'   ----------------------------------------------------------------------

'   MsgBox(prompt[, buttons][, title][, helpfile, context])

'   Function used to get input string from user using GUI interface. For
'   this purpose it creates separate window, where user can input a text.

'ignore result
MsgBox "Message text", vbOkOnly, "Message title"
'check user answer
ans = MsgBox("Accept or decline", vbOkCancel, "Please, choose!")
if asw = vbOk then
    res = "accepted"
else
    res = "declined"
end if
WScript.Echo "user was " & res & " your offer!"


'   3. Input Box
'   ----------------------------------------------------------------------

'   InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])

'   Get input string from user using separate GUI-window.

' invoke input window
input = InputBox("Type a value", "Input Box", "0")
' display result in message box window
MsgBox "User was entered '" & input & "' into the input box."

'   4. Standard input/output streams
'   ----------------------------------------------------------------------

'   Standard input output and error streams available as poroperties of
'   `WScript` object if script was invoked by means of _cscript.exe_
'   interpreter. Properties names are: `StdIn`, `StdOut` and `StdErr`
'   respectively. 

'force CR + LF at the end of the string
WScript.StdOut.WriteLine "input something here and press enter:"
'read line from standard input and return to execution
line = WScript.StdIn.ReadLine()
'write string to standard output without inserting CR + LF
WScript.StdOut.Write "'" & line
WScript.StdOut.Write "' was read!" & vbCRLF  'do line end explicitly
'output to error stream
WScript.StdErr.WriteLine "here can be an error message"
do 'per character read from StdIn
    ch = WScript.StdIn.Read(1) 'here stopped until user press 'Enter'
    WScript.StdOut.Write ch
    if ch = vbLF and old_ch = vbCR then exit do
    old_ch = ch
loop


'   5. Working with arguments
'   ----------------------------------------------------------------------

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
    WScript.StdOut.WriteLine "arg " & (i + 1) & " = '" & args(i) & "'"
next

