'here working with dates and time
    
sub echo(msg)
    wscript.echo msg
end sub

echo Date()
echo Time()
echo Now()

tomorrow = Date() + 1

echo "Tomorrow is: " & DateValue(tomorrow)

echo "prev month last day is: " & Date() - Day(Date())
