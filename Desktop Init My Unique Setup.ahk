SetWorkingDir, C:\Users\%username%\Desktop\Scripts

Browser = "C:\Program Files\Google\Chrome\Application\chrome.exe"

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

Loop, 20
{
Send, #^{Left}
}


; Configure Desktop #1   ; Main6
Run "C:\Users\%username%\Desktop\Phone intake.txt"
Run %Browser% "https://apps.8x8.com" "https://analytics.8x8.com/shared-dashboard/ef9816e8-9f2f-4b1e-923d-f523a50284fa/light"
Send, #1  ; Outlook
Send, #2  ; Teams
Send, #3  ; 8x8
Send, #6  ; Open CWM
Loop
{
if (A_Index > 60)
        break  ; Terminate the loop
ToolTip, %A_Index%
sleep 1000
}
ToolTip


; Configure Desktop #2
Send, #^{Right}
sleep 200
Run %Browser% "https://na1vsa40.kaseya.net/Authenticate/?sso=true&firstIn=true#navigation:2001" "https://hd123.me/Login?Reason=0"
sleep 2000
Click, 95, 59
sleep 3000
Click, 937, 711
sleep 3000
Click, 896, 519
sleep 500


; Configure Desktop #3
Send, #^{Right}
sleep 200
Run %Browser% "http://google.com"
sleep 500
Run "C:\Users\%username%\Documents\SciTE4AHK_v3.1.0_Portable\SciTE\SciTE.exe"
sleep 15000  ; SciTE takes a moment to load if it tries to open multiple files


; Configure Desktop #4
Send, #^{Right}
sleep 200
Run %Browser% "http://google.com"
sleep 500

; Configure Desktop #5
Send, #^{Right}
sleep 200
Run %Browser% "http://google.com"
sleep 500

; Configure Desktop #6
Send, #^{Right}
sleep 200
sleep 500

; Configure Desktop #7
Send, #^{Right}
sleep 200

; Configure Desktop #8
Send, #^{Right}
sleep 200

; Configure Desktop #9
Send, #^{Right}
Run "C:\Users\%username%\Desktop\Scripts"
sleep 500


ExitApp
