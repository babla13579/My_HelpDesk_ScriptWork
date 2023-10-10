# My_HelpDesk_ScriptWork
Broken due to me trimming out alot of company specific stuff... I haven't even tried running this after I EXTENSIVELY trimmed it down.
Sharing for the overall script template  and  some useful functions if you write custom scriptz

I wrote this script so multiple people can speed up their jobs (IT Helpdesk)
I designed the script to have a user customizable config file --- It pulls sensible defaults from a template config file.
Grouped many of the keyshortcuts to "fall under" a function key... This was inspired by control characters in TMUX and GNU Screen

Also, I wrote my AHK script in "SciTE4Autohotkey"... You'll see alot of ;{ and ;} for code folding...

Below is what I think is the most useful stuff that people could modify and impliment in their environment/scripts pretty easily:
Some (not all) of this is written by me...

Demo-AHK.ahk --- Overall Hotkey Layout
    As mentioned above, I designed my script to be primarially based on grouping hotkeys to make the keycombos overallless difficult to remember
    For example
        F1:: return ; zzz_ Control Character for shortcuts - Custom Clipboard, Desktop Management
        {
        #If (A_PriorHotKey = "F1" AND A_TimeSincePriorHotkey < 3000)
          ; Define your "sub F1" hotkeys here... 
          ; when running the script when you push the F1 key, you have "A_TimeSincePriorHotkey" milliseconds to push a sub hotkey
          ; In my opinion, formatting your hotkeys by "job function" to fall under a specific FN key helps to make remembering things less of a burden
          ; Plus this kind of format greatly increases the total number of potential hotkeys that you can use
        #If

Demo-AHK.ahk     ^!+(::     Syntax Checking   (coded by me)
    Does a quick count of Parens and Curly Brackets.

Demo-AHK.ahk     !`::       Quick AHK Syntax lookup     (coded by someone else)
     Opens AHK Help Docs  and searches for the text under the blinking cursor (I think)
     This does a very good job to quickly lookup commands when you dont remember the syntax quickly.

Demo-AHK.ahk     ^!m::     Syncs M: Drive Scripts to Local Desktop      (coded by me)
     I did all my coding work on a shared/mapped windows folder.
         See "Update_Scripts.bat"  --- multithreaded robocoby was used for fast directory syncronization
         For hardcore code development, this aint quite as good as a code management platform like Github / SVN... but it was good enough for my needs. Plus I dont risk putting sensitive info on the internet.
     However, I had the users (and me) run my code from their local C: drive... As the C: drive is always accessable. 
     This hotkey allowed users (and me) to quickly sync the M: drive to the local C: drive to update the script (and other files) to include the newest features.

Demo-AHK.ahk     ^!+m::   Backup: Creates Backup of Local (C: drive) Scripts     (coded by me)
    Used Powershell script to create a .Zip of the current C: drive scripts.
    This was useful for backup purposes  AND  continued development when the shared/mapped windows folder that I normally edit code from isn't reachable on the network.
        I could develop code and push my updates to the shared folder whenever I regained network access 

Demo-AHK.ahk     reload_Function      (coded by me)
    Flashes screen black so the user knows the script was reloaded...
    Also has a 50 millisecond sleep timer so multiple instances of the script dont launch
    
Demo-AHK.ahk     ^!/::Run, C:\Users\%username%\Pictures\ScreenShots\%A_YYYY%\%A_MM% ; GreenShot Screenshots  Opens Screenshot folder GreenShot software - Output FileName Pattern ${YYYY}\${MM}\zz_${DD}_${hh}${mm}_${NUM}
    Nothing special... Including this as I HIGHLY reccomend GreenShot to quickly take screenshots.
    This shortcut opened my screenshot folder for the month...

Demo-AHK.ahk     ^!p::    Quickly switched to the phone call tab in google chrome... Useful when I gotta answer a phone call quick.

Demo-AHK.ahk --- INI User Configurable Settings          https://www.reddit.com/r/AutoHotkey/comments/s1it4j/automagically_readwrite_configuration_files/
    See lines 4-13,103
    See functions ---  Reset_INI_File_From_Template   and   Update_INI_File_From_Template    (coded by me)
    See the {F9} hotkey section --- this section is dedicated to adjusting / reseting the user / template INI files
        Note:  Update_INI_File_From_Template
                   Compares the template INI file to the User's INI file.
                   If the users INI file doesn't have all the key-values pairs contained in the template, the missing key-value pairs are added to the user config file.
                   Does NOT delete key-value pairs from user config if that key-value pair is removed from template
    
Demo-AHK.ahk --- DisplayTooltip   function     (coded by me)
    Used to briefly display a tooltip next to the cursor for a specified millisecond time...

Demo-AHK.ahk --- "Click Tracking Functions"     (coded by me)
    These functions mark/recall clicks relative to the upper left corner of the active window...
    Useful to quickly gather physical click cordinates
        !+LButton::    Adds current click cordinates to variable.
        !+RButton::    Recalls list of click cordinates from variable
            Reload script to reset

Demo-AHK.ahk ---  ^WheelUp::    and    ^WheelDown::
    Copies Username or Password of last looked up credential from Windows Credential manager.
        See {F3} section
    This was heavily used by me and my team as I built in additional logic (not shown in this demo script) to instantly lookup admin credentials for all the companies that we supported.
        Each tech easily touched 30-60 PCs a day so this feature save alot of time

Demo-AHK.ahk --- Reliable Google Chrome Tab Switching    https://www.youtube.com/watch?v=GYHOtC737Ag
    See functions ---  switchChromeTabByTitle    and    getChromeTabTitle

Demo-AHK.ahk --- Virtual Desktop Management              https://www.computerhope.com/tips/tip224.htm
    My code is tested and works on both Windows 10 and 11.
    See Auto-Execute Section    search for   ;{ Virtual Desktop Management
    See lines  303-406
    Search for ::switchDesktopByNumber  to see where this is implemented.
        The Virtual Desktop Management is nice so you can run specific applications on whatever desktop.
        For instance, I have Desktop 1 as my "Communication Desktop"... Outlook, MS Teams, Ticketing System, Soft Phone.
                             Desktop 2 was for connecting to remote computers
                             Desktop 3 was for Developer / Coding work
                             etc.

Other stuff to take a look at that will require more effort on your part...

    The {F4} section:
        This section is dedicated to my Rufaydium work (Web Browsing Automation)
        Probably none of this will work on your computer... But it can give you an idea of what is possible.  I'm leaving most of my work in case you want to review.
        Also, be aware that your HDD can fill up with temp files from your browser if rufaydium doesn't exit nicely... Not certain why this occurs (root problem)
            See  C:\Users\%username%\AppData\Local\Temp\scoped_dir*  and  C:\Users\%username%\AppData\Local\Temp\edge_BITS_*
        My first Rufaydium project that I worked on was loved by my team as it saved hours of manual labor...  See   u:: ; {F4} -- Unknown Ticket Automation
            This went to our ticketing system and looped over a particular queue and (for each ticket) checked to see if it could be assigned/handled via my script.
            Basically, I got the interesting ticket details EmailSenderAddress, TicketTitle, TicketDetails  and  I concatinated those values and called a lookup function "Ticket_AccountLookup" to attempt to lookup the appropriate info to auto-handle the ticket
            
