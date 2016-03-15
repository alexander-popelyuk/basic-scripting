'declare an array
dim val(23) '<<< here '23' is max index, not item count!!

'fill array with random values
for i = 0 to 23 step 1
   val(i) = Int(rnd * 100)
next

'print array content
wscript.echo "static array content:"
c = 0
for each x in val
    c = c + 1
    msg = c & ". next array value is: " & x
    if c < 10 then
        msg = "0" & msg
    end if
    wscript.echo msg
next

'define dynamic array
dim v8()

'set  array size
redim v8(9)

'fill array with random values
for i = 0 to 9 step 1
   v8(i) = Int(rnd * 100)
next

'print array data
wscript.echo "dynamic array, size 10:"
c = 0
for each x in v8
    c = c + 1
    msg = c & ". next array value is: " & x
    if c < 10 then
        msg = "0" & msg
    end if
    wscript.echo msg
next

'change array size
redim preserve v8(4)

'print new array
wscript.echo "dynamic array, size 5:"
c = 0
for each x in v8
    c = c + 1
    msg = c & ". next array value is: " & x
    if c < 10 then
        msg = "0" & msg
    end if
    wscript.echo msg
next

