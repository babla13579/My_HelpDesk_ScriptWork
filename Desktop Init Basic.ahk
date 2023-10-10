; Close all open desktops
Loop, 20
{
Send, #^{F4}
}

; Open 8 more desktops for a total of 9
Loop, 8
{
Send, #^d
}

; Go back to first desktop
Loop, 20
{
Send, #^{Left}
}

Run "C:\Users\%username%\Desktop\Scripts"
ExitApp
