on error resume next 'enable error suppressing
lajfs
failed = err.number <> 0
err_cause = err.description
on error goto 0 'disable error suppressing

'if something was detected, print it
if failed then
    MsgBox "the folowing problem occured: " & err_cause
end if
