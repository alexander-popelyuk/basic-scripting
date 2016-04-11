'----------------------------------------------------------------------
' error handling
'----------------------------------------------------------------------

' VBScript provide for error handling two things:
'  1. statements 'on error resume next' and 'on error goto 0' to
'     enable/disable processing next statement after error occur;
'  2. 'Err' global object, which provide information about errors.

' 'Err' object has 3 properties:
'  1. 'Number' - code of the occurred error (zero - if no errors);
'  2. 'Description' - descriptive message about a error;
'  3. 'Source' - provide a source object name to determine the error
'     source.
' 'Err' object also provide two methods to handle error:
'  1. 'Raise' - rises a error with specified parameters;
'  2. 'Clear' - clears a last occurred error, need just to explicitly
'     clear errors because it cleared automatically after 'on error
'     resume next' statement.

'error processing procedure
sub CheckErrors()
    if Err.Number <> 0 then
        WScript.Echo ""
        WScript.Echo "A error occurred!"
        WScript.Echo "Number: " & Err.Number
        WScript.Echo "Description: " & Err.Description
        WScript.Echo "Source: " & Err.Source
        Err.Clear
    end if
end sub

on error resume next

'run erroneous code
k = GetGoodValue()
CheckErrors()

'explicitly raise a error
Err.Raise vbObjectError + 29292, "My own error", "Test error raising!"
CheckErrors()

on error goto 0

