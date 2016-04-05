sub testV(ByVal x)
    echo "testV(): before x(0) = " & x(0)
    x(0) = x(0) + 1
    echo "testV(): after  x(0) = " & x(0)
end sub

sub testR(ByRef x)
    echo "testR(): before x(0) = " & x(0)
    x(0) = x(0) + 1
    echo "testR(): after  x(0) = " & x(0)
end sub

sub echo(x)
    wscript.echo x
end sub

dim l(9)
l(0) = 1

echo "before calls l = " & l(0)
testV(l)
echo "after testV() l = " & l(0)
testR(l)
echo "after testR() l = " & l(0)
call testR(l)
echo "after testR() call l = " & l(0)
testR l
echo "after testR l = " & l(0)

