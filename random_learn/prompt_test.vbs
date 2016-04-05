'prompt user for login and password
login = InputBox("Please enter the login:", "Authorization", "user")

'check user is register in system
select case login
    case "user", "admin"
        pass = InputBox("Please enter password for user '" & _
            login & "':", "Authorization")
    case else
        MsgBox "Unknown user '" & login & "', access denied!", _
            vbOkOnly + vbCritical, "Authorization failed!"
        wscript.quit 1
end select

'check password for the user
if login = "user" and pass = "simple pass" then 
    MsgBox "Hello user, access granted!", _
        vbOkOnly + vbInformation, "Authorization completed!"
elseif login = "admin" and pass = "complicated password" then
    MsgBox "Hello admin, access granted!", _
        vbOkOnly + vbInformation, "Authorization completed!"
else
    MsgBox "Incorrect password, access denied!", _
        vbOkOnly + vbCritical, "Authorization failed!"
    wscript.quit 1
end if

