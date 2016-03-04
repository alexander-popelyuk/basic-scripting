'k = 4
'wscript.echo "something here, k =", k
'
'wscript.echo "program got", wscript.arguments.length, "argument(s)"
'for i = 0 to wscript.arguments.length - 1
'    wscript.echo "arg", i+1, "=", wscript.arguments(i)
'next

if Hour(Time()) < 12 then
    wscript.stdout.writeLine("it's morning!")
else
    for i = 0 to 20
        wscript.stdout.writeLine("it's evening!")
    next
    select case Hour(Time())
        case 22:
            wscript.echo "new item"
        case 99:
            wscript.echo "got this shit"
        case else:
            wscript.echo "got else"
    end select
end if

t = 5

do while t
    wscript.echo "inside while... t =", t
    t = t - 1
loop

wscript.echo ""

do until t > 5
    wscript.echo "loop 2, t =", t
    t = t + 1
loop

wscript.echo ""

do
    wscript.echo "loop 3, t =", t
    t = t - 1
loop while t

wscript.echo ""

do 
    wscript.echo "loop 4, t =", t
    t = t + 1
loop until t > 5

wscript.echo ""

do
    wscript.echo "inside simple loop, t =", t
    t = t - 1
    if (t = 0) then
        exit do
    end if
loop

v = 3
wscript.echo "v =", v, "-v =", (not v) + 1
for v = 99 to 0 step -2
    wscript.echo "v =", v
next
