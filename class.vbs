'test class creation

'print argument
private sub echo(x)
    wscript.echo x
end sub

class test

    'constructor
    private sub Class_Initialize()
        echo "constructor was called!"
    end sub

    'destructor
    private sub Class_Terminate()
        echo "destructor was called!"
    end sub

    'test method1
    sub method1()
        echo "method1 was called"
    end sub

end class

echo "script started"

'create test class
set t = new test

t.method1

set t = Nothing

echo "end of script"
