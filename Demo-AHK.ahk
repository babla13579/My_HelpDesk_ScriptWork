;{  Auto-Execute Section    https://www.autohotkey.com/docs/v1/Scripts.htm#auto
SetWorkingDir, C:\Users\%username%\Desktop\Scripts\
;{  Initialize / Update / Read User Config File
User_Config_File_Path = C:\Users\%username%\Documents\MY_COMPANY_NAME_Script_User_Config_%username%.ini
User_Config_Template_Path = C:\Users\%username%\Desktop\Scripts\Reference_Files\File_Templates\MY_COMPANY_NAME_Script_User_Config_Template.ini
MDrive_User_Config_Template_Path = M:\Shared_SMB\Folder_Path\Reference_Files\File_Templates\MY_COMPANY_NAME_Script_User_Config_Template.ini
IfNotExist, %User_Config_File_Path%
{
  Reset_INI_File_From_Template(User_Config_File_Path,User_Config_Template_Path)
} else {
  Update_INI_File_From_Template(User_Config_File_Path,User_Config_Template_Path)
}
User_Config_File := Ini(User_Config_File_Path)
;}
;{  Ask User To Change Defaults --- User Config File
  if (User_Config_File.TicketingSystem.Full_Name = "CHANGEME_FullNameAsAppearsInTicketingSystem") {
    InputBox, UserInput, User Configuration File Adjustment, Please enter your Full Name as it appears in TicketingSystem
    if (UserInput) {
      User_Config_File.TicketingSystem.Full_Name :=UserInput
    }
  }
  if (User_Config_File.Win_Cred_Man.User_EmployeeTimeTrackingSoftware = "CHANGEME_EmployeeTimeTrackingSoftwareusername") {
    InputBox, UserInput, User Configuration File Adjustment - Windows Credential Manager, Please enter your EmployeeTimeTrackingSoftware username
    if (UserInput) {
      User_Config_File.Win_Cred_Man.User_EmployeeTimeTrackingSoftware :=UserInput
      InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your EmployeeTimeTrackingSoftware password (hidden input), hide
      if (UserInput1) {
        CredWrite("EmployeeTimeTrackingSoftware_cred",User_Config_File.Win_Cred_Man.User_EmployeeTimeTrackingSoftware,UserInput1)
      }
    }
  }
  if (User_Config_File.Win_Cred_Man.User_AD = "CHANGEME_username") {
    UserInput=%username%
    fileread, text, %User_Config_File_Path%
    newtext := strreplace(text, "CHANGEME_username", UserInput)
    filedelete, %User_Config_File_Path%
    sleep 20
    fileappend, %newtext%, %User_Config_File_Path%
    User_Config_File := Ini(User_Config_File_Path)
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your MY_COMPANY_NAME Active Directory Password (hidden input), hide
    if (UserInput1) {
      CredWrite("MY_COMPANY_NAME_AD_cred",User_Config_File.Win_Cred_Man.User_AD,UserInput1)
    }
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your company123.Admin password (hidden input), hide
    if (UserInput1) {
      CredWrite("company123_Admin_cred",User_Config_File.Win_Cred_Man.User_company123_Admin,UserInput1)
    }
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your Office 365 password (hidden input), hide
    if (UserInput1) {
      CredWrite("Office365_cred",User_Config_File.Win_Cred_Man.User_O365,UserInput1)
    }
    reload_Function()   ; Reload Script
  }
;}
if (User_Config_File.Win_Cred_Man.Update_User_Passwords_Upon_Reload) {
;{ EmployeeTimeTrackingSoftware
  MsgBox, 292, Change - EmployeeTimeTrackingSoftware Password, Update your EmployeeTimeTrackingSoftware Password?
  IfMsgBox Yes
  {
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your EmployeeTimeTrackingSoftware password (hidden input), hide
    if (UserInput1) {
      CredWrite("EmployeeTimeTrackingSoftware_cred",User_Config_File.Win_Cred_Man.User_EmployeeTimeTrackingSoftware,UserInput1)
    }
  }
;}
;{ MY_COMPANY_NAME AD
  MsgBox, 292, Change - Local Domain Password, Update your MY_COMPANY_NAME AD Password?
  IfMsgBox Yes
  {
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your MY_COMPANY_NAME Active Directory Password (hidden input), hide
    if (UserInput1) {
      CredWrite("MY_COMPANY_NAME_AD_cred",User_Config_File.Win_Cred_Man.User_AD,UserInput1)
    }
  }
;}
;{ company123 Admin
  MsgBox, 292, Change - company123.Admin Password, Update your company123.Admin Password?
  IfMsgBox Yes
  {
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your company123.Admin password (hidden input), hide
    if (UserInput1) {
      CredWrite("company123_Admin_cred",User_Config_File.Win_Cred_Man.User_company123_Admin,UserInput1)
    }
  }
;}
;{ Office 365
  MsgBox, 292, Change - Office 365 Password, Update your Office 365 Password?
  IfMsgBox Yes
  {
    InputBox, UserInput1, User Configuration File Adjustment - Windows Credential Manager, Please enter your Office 365 password (hidden input), hide
    if (UserInput1) {
      CredWrite("Office365_cred",User_Config_File.Win_Cred_Man.User_O365,UserInput1)
    }
  }
;}

  User_Config_File.Win_Cred_Man.Update_User_Passwords_Upon_Reload := 0
  DisplayTooltip("Passwords Updated",4000,1)
}
;DetectHiddenWindows, On
#InstallKeybdHook
;{  #INCLUDE
#Include AutoHotKey-Main\INI\Ini.ahk
#Include AutoHotKey-Main\Rufaydium\Rufaydium.ahk
;}
Kaseya_Browser := User_Config_File.Browser.Kaseya
Office_Browser := User_Config_File.Browser.Office_Portal
;~ MsgBox % Kaseya_Browser . "`n" . Office_Browser
;{   Launch  HotKey_List.ahk   for the #{F1} hotkey
Run cmd.exe /c start "" "C:\Users\%username%\Desktop\Scripts\AutoHotKey-Main\HotKey_List.ahk",, hide
;}
;{ Virtual Desktop Management
DesktopCount = 2 ; Windows starts with 2 desktops at boot
CurrentDesktop = 1 ; Desktop count is 1-indexed (Microsoft numbers them this way)
SetKeyDelay, 100   ; was 75
mapDesktopsFromRegistry()
;}
;}

;{  ALL Functions and HotStrings
; Windows + z  ---- SUSPEND/RESUME THIS AHK SCRIPT... if default windows button bindings are needed instead of AHK bindings
#z::Suspend, Toggle  ; aaa_  Suspend/Resume Hotkeys - Restores Windows Keyshortcuts


;{  Kaseya Lookup Functions
VSA_JSON_Manipulate(agentguid,connectionmode) {
  FileRead, JSON_Template, C:\Users\%username%\Desktop\Scripts\Reference_Files\File_Templates\JSON_Template_File.txt
  StringReplace, JSON_Current, JSON_Template, __CHANGE_ME__, %agentguid%
  StringReplace, JSON_Current, JSON_Current, __CHANGE_ME_TOO__, %connectionmode%
  return %JSON_Current%
}
Ticket_AccountLookupTicket_AccountLookup(All_Ticket_Text) {
    if ( InStr(All_Ticket_Text, " example_uniq123") || InStr(All_Ticket_Text, "example_uniq123aaa") ) { ; Company A
      ; " example_uniq123"     Special RMM Tool
      ; example_uniq123aaa    Common Email
      account = Company A
    } else if (  RegExMatch(All_Ticket_Text, "i).*hc-.*\b(?:w10|w11).*(?:ws|nb).*") || InStr(All_Ticket_Text, "Comapny B") ) { ; Company B
      ; "i)^.*hc-.*\b(?:w10|w11).*(?:ws|nb).*$"    HC-*[W10|W11]*[WS|NB]   Matches PC Name  case insensitive
      ; Comapny B
      account = Company B
    } else { ; No match found
      uticket_notmatched_count := uticket_notmatched_count + 1     ;  Increment counter
    }
    ;~ msgbox % "." . account . "."
    if (!account) {
      ;~ msgbox, %All_Ticket_Text%   ;
      if (account = "My_Company") {  ; Dont trust this... To many other companies have alerts that include My_Company in the ticket details
        account =  ;Null
      }
    }
    if (account) {
      return account
    }
}
;}


;{  NON -- Kaseya Lookup Functions   ---  Script Dependencies

;{  reload_Function    Flash_Screen    DisplayToolTip
reload_Function() {    ; Refresh HotKey help list    Flash Screen    Reload
  ;~ Run cmd.exe /c taskkill /f /t /im:msedgedriver.exe , , Min    ; Kills the web driver  /t flag  also kills open browsers launched from webdriver
  flash_Screen("Black")
  sleep 50  ; Sleep needed to prevent multiple script instances in case someone holds down the Reload hotkey
  Reload   ; Reloads script
}
flash_Screen(f_Color) {
  ; Used to flash screen
  Gui,Color,%f_Color%
  Gui,+AlwaysOnTop +ToolWindow -Caption
  Gui,Show,% "w" A_ScreenWidth "." "h" A_ScreenHeight + 100 "." "Hide NA"

  Gui,Show
  Sleep,50
  Gui,Hide
}
flash_Screen_Rainbow() { ; Rainbow = Cool

  Colors=Red|Orange|Yellow|Green|Blue|Indigo|Violet
  Loop, parse, Colors, `|
  {
    ; Used to flash screen
    Gui,Color,%A_LoopField%
    Gui,+AlwaysOnTop +ToolWindow -Caption
    Gui,Show,% "w" A_ScreenWidth "." "h" A_ScreenHeight + 100 "." "Hide NA"

    Gui,Show
    Sleep,50
    Gui,Hide
  }
}
DisplayTooltip(msg_contents,TT_timer,TT_num) {
{
  #Persistent
  ToolTip, %msg_contents%,,,%TT_num%
  SetTimer, RemoveToolTip, -%TT_timer%
  return
}
RemoveToolTip:
  ToolTip,,,,%TT_num%
  return
}
;}
;{  User Config File    Reset  /  Ensure All Keys Exist
Reset_INI_File_From_Template(INI_File_Path, INI_Template_Path) {
  run, cmd.exe /c copy /y %INI_Template_Path% %INI_File_Path%,,hide,PID
  Process, WaitClose, %PID%
  return
}
Update_INI_File_From_Template(INI_File_Path,Template_INI_File_Path) {  # loops over template INI file. If any user runs script and doesn't have all keys. This function will add missing keys
  Loop, read, %Template_INI_File_Path%
  {
    if RegExMatch(A_LoopReadLine, "^\[.*\]$") {   ;   Matches any line that starts with  [  and ends with  ]     INI Sections start and end with square brackets [Example]
      Cur_Sec := StrReplace(A_LoopReadLine,"[")
      Cur_Sec := StrReplace(Cur_Sec,"]")
    } else {
      if (Cur_Sec) {  ; Only check if we are in an .ini SECTION
        if RegExMatch(A_LoopReadLine, "^[-A-Za-z0-9_]+=.*$") {   ; Key Match
          StringSplit, kv_pair, A_LoopReadLine , =
          if ( (kv_pair1 != "") && (kv_pair2 != "") ) {   ; Key Value Pair is Defined in User_Config_Template
            Found=0 ; Prep for loop
            Loop, read, %INI_File_Path%
            {
              if RegExMatch(A_LoopReadLine, "^\[.*\]$") {  ; Matches any INI Sections
                S := StrReplace(A_LoopReadLine,"[")
                S := StrReplace(S,"]")
              } else {
                if (S) {  ; Only check if we are in an .ini SECTION
                  if RegExMatch(A_LoopReadLine, "^[-A-Za-z0-9_]+=.*$") {   ; Key Match
                    StringSplit, S_keyval_pair, A_LoopReadLine , =
                    if ( (S_keyval_pair1 != "") && (S_keyval_pair2 != "") ) {   ; Key Value Pair is Defined in User_Config_File
                      if ( (Cur_Sec = S) && (S_keyval_pair1 = kv_pair1) ) {  ; Match
                        Found=1  ; True
                        break
                      }
                    }
                  }
                }
              }
            }
            if (Found) {
              continue
            } else {
              IniWrite, %kv_pair2%, %INI_File_Path%, %Cur_Sec%, %kv_pair1%
            }
          }
        }
      }
    }
  }
}
;}
;{  Dependencies  --- Functions
;{ Array Searching
HasVal(haystack, needle) {
    for index, value in haystack
        if (value = needle)
            return index
    if !(IsObject(haystack))
        throw Exception("Bad haystack!", -1, haystack)
    return 0
}
;}
;{  Base 64  Encode  /  Decode
b64Encode(string)
{
    VarSetCapacity(bin, StrPut(string, "UTF-8")) && len := StrPut(string, &bin, "UTF-8") - 1
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x1, "ptr", 0, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    VarSetCapacity(buf, size << 1, 0)
    if !(DllCall("crypt32\CryptBinaryToString", "ptr", &bin, "uint", len, "uint", 0x1, "ptr", &buf, "uint*", size))
        throw Exception("CryptBinaryToString failed", -1)
    return StrGet(&buf)
}
b64Decode(string)
{
    if !(DllCall("crypt32\CryptStringToBinary", "ptr", &string, "uint", 0, "uint", 0x1, "ptr", 0, "uint*", size, "ptr", 0, "ptr", 0))
        throw Exception("CryptStringToBinary failed", -1)
    VarSetCapacity(buf, size, 0)
    if !(DllCall("crypt32\CryptStringToBinary", "ptr", &string, "uint", 0, "uint", 0x1, "ptr", &buf, "uint*", size, "ptr", 0, "ptr", 0))
        throw Exception("CryptStringToBinary failed", -1)
    return StrGet(&buf, size, "UTF-8")
}
;}
;{  Chrome Tabs   Switching by Title - Get Title - Get URL
getChromeTabTitle()
{
  WinGetTitle, Title, A
  return %Title%
}
getChromeTabURL()
{
  prev_clip = %clipboard%
  Send, !d
  Send, ^c
  Send, {Esc}
  URL = %clipboard%
  clipboard = %prev_clip%
  return %URL%
}
switchChromeTabByTitle(tabName) {
  SetTitleMatchMode, 2
  WinGet, ChromeCount, Count, ahk_exe chrome.exe
  activeTab := ""
  firstTab := ""

  Loop %ChromeCount%
  {
    WinActivateBottom, ahk_exe chrome.exe
    WinWaitActive, ahk_exe chrome.exe
    WinGetTitle, firstTab, A
    If InStr(firstTab, tabName)
      break
    While !(activeTab = firstTab)
    {
      Send ^{tab}
      Sleep 50
      WinGetTitle, activeTab, A
      If InStr(activeTab, tabName)
        break 2 ;https://www.autohotkey.com/docs/commands/Break.htm
    }
  }
  return
}
;}
;{  Desktop Management
mapDesktopsFromRegistry() {
; This function examines the registry to build an accurate list of the current virtual desktops and which one we're currently on.
; Current desktop UUID appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\1\VirtualDesktops
; List of desktops appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops
;
 global CurrentDesktop, DesktopCount
 ; Get the current desktop UUID. Length should be 32 always, but there's no guarantee this couldn't change in a later Windows release so we check.
 IdLength := 32
 SessionId := getSessionId()
 if (SessionId) {
 RegRead, CurrentDesktopId, HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\%SessionId%\VirtualDesktops, CurrentVirtualDesktop
 if (CurrentDesktopId) {
 IdLength := StrLen(CurrentDesktopId)
 }
 }
 ; Get a list of the UUIDs for all virtual desktops on the system
 RegRead, DesktopList, HKEY_CURRENT_USER, SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops, VirtualDesktopIDs
 if (DesktopList) {
 DesktopListLength := StrLen(DesktopList)
 ; Figure out how many virtual desktops there are
 DesktopCount := DesktopListLength / IdLength
 }
 else {
 DesktopCount := 1
 }
 ; Parse the REG_DATA string that stores the array of UUID's for virtual desktops in the registry.
 i := 0
 while (CurrentDesktopId and i < DesktopCount) {
 StartPos := (i * IdLength) + 1
 DesktopIter := SubStr(DesktopList, StartPos, IdLength)
 OutputDebug, The iterator is pointing at %DesktopIter% and count is %i%.
 ; Break out if we find a match in the list. If we didn't find anything, keep the
 ; old guess and pray we're still correct :-D.
 if (DesktopIter = CurrentDesktopId) {
 CurrentDesktop := i + 1
 OutputDebug, Current desktop number is %CurrentDesktop% with an ID of %DesktopIter%.
 break
 }
 i++
 }
}
getSessionId()
{
; This functions finds out ID of current session.
 ProcessId := DllCall("GetCurrentProcessId", "UInt")
 if ErrorLevel {
 OutputDebug, Error getting current process id: %ErrorLevel%
 return
 }
 OutputDebug, Current Process Id: %ProcessId%
 DllCall("ProcessIdToSessionId", "UInt", ProcessId, "UInt*", SessionId)
 if ErrorLevel {
 OutputDebug, Error getting session id: %ErrorLevel%
 return
 }
 OutputDebug, Current Session Id: %SessionId%
 return SessionId
}
switchDesktopByNumber(targetDesktop)
{
  ; This function switches to the desktop number provided.
 global CurrentDesktop, DesktopCount
 ; Re-generate the list of desktops and where we fit in that. We do this because
 ; the user may have switched desktops via some other means than the script.
 mapDesktopsFromRegistry()
 ; Don't attempt to switch to an invalid desktop
 if (targetDesktop > DesktopCount || targetDesktop < 1) {
 OutputDebug, [invalid] target: %targetDesktop% current: %CurrentDesktop%
 return
 }
 WinActivate, ahk_class Shell_TrayWnd  ; Fixes the issue of active windows in intermediate desktops capturing the switch shortcut
 ; Go right until we reach the desktop we want
 while(CurrentDesktop < targetDesktop) {
 Send ^#{Right}
 CurrentDesktop++
 OutputDebug, [right] target: %targetDesktop% current: %CurrentDesktop%
 }
 ; Go left until we reach the desktop we want
 while(CurrentDesktop > targetDesktop) {
 Send ^#{Left}
 CurrentDesktop--
 OutputDebug, [left] target: %targetDesktop% current: %CurrentDesktop%
 }
}
createVirtualDesktop()
{
  ; This function creates a new virtual desktop and switches to it
 global CurrentDesktop, DesktopCount
 Send, #^d
 DesktopCount++
 CurrentDesktop = %DesktopCount%
 OutputDebug, [create] desktops: %DesktopCount% current: %CurrentDesktop%
}
deleteVirtualDesktop()
{
; This function deletes the current virtual desktop
 global CurrentDesktop, DesktopCount
 Send, #^{F4}
 DesktopCount--
 CurrentDesktop--
 OutputDebug, [delete] desktops: %DesktopCount% current: %CurrentDesktop%
}
;}
;{  Input ComboBox
InputComboBox(AuswahlListe, Hinweis := "Bitte wählen:", Titel := "", Owner := "") {
   Local
      Selection := ""
      Cancelled := 0
      Gui, InputComboBox: +hwndHGUI -MinimizeBox +Owner%Owner%
      Gui, InputComboBox: Margin, 10, 10
      Gui, InputComboBox: Add, Text, w500 h20, %Hinweis%
      Gui, InputComboBox: Add, Combobox, xm y+5 w500 gCbAutoComplete hwndHCBB, %AuswahlListe%
      Gui, InputComboBox: Add, Button, xm hwndHBT2 gInputComboBoxBtnCancel, Cancel
      GuiControlGet, P1, Pos, %HBT1%
      Gui, InputComboBox: Add, Button, x+m Default hwndHBT1 gInputComboBoxBtnOK, OK
      GuiControlGet, P2, Pos, %HBT2%
      GuiControl, Move, %HBT1%, % "w" . P2W
      GuiControl, Move, %HBT2%, % "x" . (310 - P2W)
      Gui, InputComboBox: Show, w520 h110 , % (Titel = "" ? A_ScriptName : Titel)
      WinWaitClose, ahk_id %HGUI%
      ErrorLevel := Cancelled
   Return ErrorLevel>0 ?: Selection

   InputComboBoxGuiEscape:
   InputComboBoxGuiClose:
   InputComboBoxBtnCancel:
   Gui, Destroy
   Return

   InputComboBoxBtnOK:
      GuiControlGet, Selection, , %HCBB%
      Gui, Destroy
      Cancelled := (A_ThisLabel = "InputComboBoxBtnOK" ? 0 : 1)
  Return
}
CbAutoComplete(HCBB) {
   ; CB_GETEDITSEL = 0x0140, CB_SETEDITSEL = 0x0142
   Local
   If ((GetKeyState("Delete", "P")) || (GetKeyState("Backspace", "P")))
      Return
   ; GuiControlGet, lHwnd, Hwnd, %A_GuiControl%
   SendMessage, 0x0140, 0, 0, , ahk_id %HCBB%
   Start := ErrorLevel & 0xFFFF
   GuiControlGet, CurContent, , %HCBB%, Text
   GuiControl, ChooseString, %HCBB%, %CurContent%
   If (ErrorLevel)
      GuiControl, Text, %HCBB%, %CurContent%
   Else
      GuiControlGet, CurContent, , %HCBB%
   Sel := Start | (StrLen(CurContent) << 16)
   PostMessage, 0x0142, 0, Sel, , ahk_id %HCBB%
}
;}
;{  Windows Credential Manager
CredWrite(name, username, password)
{
	VarSetCapacity(cred, 24 + A_PtrSize * 7, 0)
	cbPassword := StrLen(password)*2
	NumPut(1         , cred,  4+A_PtrSize*0, "UInt") ; Type = CRED_TYPE_GENERIC
	NumPut(&name     , cred,  8+A_PtrSize*0, "Ptr")  ; TargetName = name
	NumPut(cbPassword, cred, 16+A_PtrSize*2, "UInt") ; CredentialBlobSize
	NumPut(&password , cred, 16+A_PtrSize*3, "UInt") ; CredentialBlob
	NumPut(3         , cred, 16+A_PtrSize*4, "UInt") ; Persist = CRED_PERSIST_ENTERPRISE (roam across domain)
	NumPut(&username , cred, 24+A_PtrSize*6, "Ptr")  ; UserName
	return DllCall("Advapi32.dll\CredWriteW"
	, "Ptr", &cred ; [in] PCREDENTIALW Credential
	, "UInt", 0    ; [in] DWORD        Flags
	, "UInt") ; BOOL
}
CredDelete(name)
{
    ;~ DisplayTooltip("Deleted Cred  " . name,1500,1)
	return DllCall("Advapi32.dll\CredDeleteW"
	, "WStr", name ; [in] LPCWSTR TargetName
	, "UInt", 1    ; [in] DWORD   Type,
	, "UInt", 0    ; [in] DWORD   Flags
	, "UInt") ; BOOL
}
CredRead(name)
{
	DllCall("Advapi32.dll\CredReadW"
	, "Str", name   ; [in]  LPCWSTR      TargetName
	, "UInt", 1     ; [in]  DWORD        Type = CRED_TYPE_GENERIC (https://learn.microsoft.com/en-us/windows/win32/api/wincred/ns-wincred-credentiala)
	, "UInt", 0     ; [in]  DWORD        Flags
	, "Ptr*", pCred ; [out] PCREDENTIALW *Credential
	, "UInt") ; BOOL
	if !pCred
		return
	name := StrGet(NumGet(pCred + 8 + A_PtrSize * 0, "UPtr"), 256, "UTF-16")
	username := StrGet(NumGet(pCred + 24 + A_PtrSize * 6, "UPtr"), 256, "UTF-16")
	len := NumGet(pCred + 16 + A_PtrSize * 2, "UInt")
	password := StrGet(NumGet(pCred + 16 + A_PtrSize * 3, "UPtr"), len/2, "UTF-16")
	DllCall("Advapi32.dll\CredFree", "Ptr", pCred)
	return {"name": name, "username": username, "password": password}
}
;}
;{  Misc  Get Random Password  -  Parse Date - ColumnSwapper
GetRandPassword(length)
{
  chars=
  (Join
  a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z
  ,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z
  ,1,2,3,4,5,6,7,8,9,0
  ,!,@,#,$,`%,^,&,*,-,_,=,+,`(,`),<,>
  )
  Loop, 1
  {
    Passwords .=	""
    Loop %Length%
    {
      Sort, chars, Random D,
      Passwords .= SubStr(chars,1,1)
    }
  }
  return Passwords
}

;}
;}
;}


;{ HotStrings / TEXT EXPANSION

;{   Date  Time
:*:dt/:: ; Prints Current Date and Time
      SendInput %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%
      return
:*:dt+/:: ; Prints Current Date + XXX Days --- Prompts for Days to add
      newDate := %A_Now%
      InputBox, daystoadd, Number of days to add to current date
      newDate += %daystoadd%, days
      FormatTime, newDate,%newDate%, MM/dd/yy
      SendInput %newDate%
      return
;}
;{   Intake   Tickets / BlueStore
:*:ti/::   ; Ticket Intake
;{
      SendInput User: 
      Send, +{Enter}
      SendInput Method:  Phone
      Send, +{Enter}
      SendInput Time:  %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%
      Send, +{Enter}
      SendInput Overview: 
      Send, +{Enter}
      SendInput PC:
      return
;}
;}
:*:ntf::no trouble found  ; no trouble found
;{  Quick Notes for tickets
:*:pro/::Please reach out to the MY_COMPANY_NAME help desk (333-444-HELP) ; Please reach out to the MY_COMPANY_NAME help desk (333-444-HELP)
:*:chs/::Cleared hung session in Citrix Studio ; Cleared hung session in Citrix Studio
:*:iynh::If you require further assistance, please call into our help desk (333-444-HELP) ; If you require further assistance, please call into our help desk (333-444-HELP)
:*:iynmh::If you require further assistance, please call into our help desk (333-444-HELP) ; If you require further assistance, please call into our help desk (333-444-HELP)
:*:iyra::If you require further assistance, please call into our help desk (333-444-HELP) ; If you require further assistance, please call into our help desk (333-444-HELP)
:*:ihin::If you require further assistance, please call into our help desk (333-444-HELP) ; If you require further assistance, please call into our help desk (333-444-HELP)
:*:nusc::I have created the new user MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM based on the permissions of MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM. To receive the login credentials, please reach out to the MY_COMPANY_NAME help desk (333-444-HELP).
;}
:*:ttt:: ; Thanks, Firstname Lastname
      SendInput Thanks,
      Send, +{Enter}
      SendInput, % User_Config_File.TicketingSystem.Full_Name
      return
;}


;{  Non FUNCTION  Hotkeys

;{ Control Alt Hotkeys
;{  Reload / Sync M: Drive
^!r::
  MSEdge.QuitAllSessions() ; close all session
  MSEdge.Driver.Exit() ; then exits driver
  MSEdge_UT.QuitAllSessions()
  MSEdge_UT.Driver.Exit()
  Chrome.QuitAllSessions() ; close all session
  Chrome.Driver.Exit() ; then exits driver
  MSEdge_SingleCredential.QuitAllSessions()
  MSEdge_SingleCredential.Driver.Exit()
  reload_Function()  ; Reload Scripts - Sourced from Local Desktop
  return
Escape::  ; Reload Scripts - Sourced from Local Desktop
  send, {Esc}
  MSEdge.QuitAllSessions() ; close all session
  MSEdge.Driver.Exit() ; then exits driver
  MSEdge_UT.QuitAllSessions()
  MSEdge_UT.Driver.Exit()
  Chrome.QuitAllSessions() ; close all session
  Chrome.Driver.Exit() ; then exits driver
  MSEdge_SingleCredential.QuitAllSessions()
  MSEdge_SingleCredential.Driver.Exit()
  reload_Function()
  return
^!m::  ; aaa_  Syncs M: Drive Scripts to Local Desktop
  Run, cmd.exe /c "M:\Shared_SMB\Folder_Path\Update Scripts.bat",  , Min, M_syncpid  ; Pulls latest scripts from M: drive
  Process, WaitClose, %M_syncpid%
  flash_Screen("Purple")
  DisplayTooltip("M: Drive Has Been Synced.`n`nPress Escape to Reload Scriptz",3000,1)
  return
^!+m::  ; aaa_  Backup: Creates Backup of Local (C: drive) Scripts
  time = %A_YYYY%_%A_MM%_%A_DD%_%A_Hour%_%A_Min%_%A_Sec%
  Run, PowerShell.exe -Command Compress-Archive -Path C:\Users\%username%\Desktop\Scripts\ -CompressionLevel Optimal -DestinationPath 'C:\Users\%username%\Desktop\Script Backups\Scripts_%time%.zip',, Min, M_syncpid
  Process, WaitClose, %M_syncpid%
  flash_Screen("Purple")
  DisplayTooltip("Local Script Folder Zipped",2000,1)
  return
;}
;{  test.ahk / Folder shortcuts
^!t::Run AutoHotKey-Main\test.ahk  ; DEVELOPER USE  Runs -- .\test.ahk
;~ ^!t::flash_Screen_Rainbow()
^!+t::Run AutoHotKey-Main\test2.ahk  ; DEVELOPER USE  Runs -- .\test2.ahk
^!c::Run powershell.exe   ; Powershell
^!,::Run, C:\Users\%username%\Desktop\Scripts ; aab_  Opens local scripts folder
^!.::Run, M:\Shared_SMB\Folder_Path ; aab_  Opens M: Drive scripts folder --- Edit Scripts from this folder
^!/::Run, C:\Users\%username%\Pictures\ScreenShots\%A_YYYY%\%A_MM% ; GreenShot Screenshots  Opens Screenshot folder GreenShot software - Output FileName Pattern ${YYYY}\${MM}\zz_${DD}_${hh}${mm}_${NUM}
;}
;{  Phone - Switch to 8x8
^!p:: ; Chrome -- Switch to 8x8 tab --- Switches to TicketingSystem Desktop
    switchDesktopByNumber(User_Config_File.DesktopSwitch.TicketingSystem)
    switchChromeTabByTitle("(MY_COMPANY_NAMEc0")    ; 8x8
    return
;}
^!+(::   ; DEVELOPER USE -- Counts the number of brackets and Parens in the clipboard... Useful for syntax checking
    StringReplace clipboard,clipboard,{,{,UseErrorLevel
    LeftBracketCount = %ErrorLevel%
    StringReplace clipboard,clipboard,(,(,UseErrorLevel
    LeftParanCount = %ErrorLevel%
    StringReplace clipboard,clipboard,},},UseErrorLevel
    RightBracketCount = %ErrorLevel%
    StringReplace clipboard,clipboard,),),UseErrorLevel
    RightParanCount = %ErrorLevel%
    Msgbox % "{ --- " . LeftBracketCount . "`n} --- " . RightBracketCount . "`n( --- " . LeftParanCount . "`n) --- " . RightParanCount
    return
;}
;}

;{  Chrome Duplicate Tab / Open Clipboard in New Tab
#IfWinActive ahk_exe chrome.exe
^!w::  ; Chrome -- Opens clipboard in new tab
       Send, !d
       Send, ^v
       Send, !{Enter}
       return
^!d::  ; Chrome -- Duplicates current tab
       Send, !d
       Send, !{Enter}
       return
#IfWinActive
;}

;{ Alt Backtick --- AHK Command Help  ---- Really Nice to lookup AHK command syntax quickly
!`::  ; DEVELOPER USE -- Alt Backtick --- Opens AHK Command Help for highlighted command
SetWinDelay 10
SetKeyDelay 0
AutoTrim, On

if A_OSType = WIN32_WINDOWS  ; Windows 9x
	Sleep, 500  ; Give time for the user to release the key.

C_ClipboardPrev = %clipboard%
clipboard =
; Use the highlighted word if there is one (since sometimes the user might
; intentionally highlight something that isn't a command):
Send, ^c
ClipWait, 0.1
if ErrorLevel <> 0
{
	; Get the entire line because editors treat cursor navigation keys differently:
	Send, {home}+{end}^c
	ClipWait, 0.2
	if ErrorLevel <> 0  ; Rare, so no error is reported.
	{
		clipboard = %C_ClipboardPrev%
		return
	}
}
C_Cmd = %clipboard%  ; This will trim leading and trailing tabs & spaces.
clipboard = %C_ClipboardPrev%  ; Restore the original clipboard for the user.
Loop, parse, C_Cmd, %A_Space%`,  ; The first space or comma is the end of the command.
{
	C_Cmd = %A_LoopField%
	break ; i.e. we only need one interation.
}
IfWinNotExist, AutoHotkey Help
{
	; Determine AutoHotkey's location:
	RegRead, ahk_dir, HKEY_LOCAL_MACHINE, SOFTWARE\AutoHotkey, InstallDir
	if ErrorLevel  ; Not found, so look for it in some other common locations.
	{
		if A_AhkPath
			SplitPath, A_AhkPath,, ahk_dir
		else IfExist ..\..\AutoHotkey.chm
			ahk_dir = ..\..
		else IfExist %A_ProgramFiles%\AutoHotkey\AutoHotkey.chm
			ahk_dir = %A_ProgramFiles%\AutoHotkey
		else
		{
			MsgBox Could not find the AutoHotkey folder.
			return
		}
	}
	Run %ahk_dir%\AutoHotkey.chm
	WinWait AutoHotkey Help
}
; The above has set the "last found" window which we use below:
WinActivate
WinWaitActive
StringReplace, C_Cmd, C_Cmd, #, {#}
Send, !n{home}+{end}%C_Cmd%{enter}
return
;}
;}
;}

;{ F1 - F4     Standard Hotkeys
F1:: return ; zzz_ Control Character for shortcuts - Custom Clipboard, Desktop Management
{
#If (A_PriorHotKey = "F1" AND A_TimeSincePriorHotkey < User_Config_File.Timer.FN_Key_Timer)

;{ Virtual Desktop Switching  1-9
    1::switchDesktopByNumber(1) ; {F1} --  SwitchToDesktop num
    2::switchDesktopByNumber(2) ; {F1} --  SwitchToDesktop num
    3::switchDesktopByNumber(3) ; {F1} --  SwitchToDesktop num
    4::switchDesktopByNumber(4) ; {F1} --  SwitchToDesktop num
    5::switchDesktopByNumber(5) ; {F1} --  SwitchToDesktop num
    6::switchDesktopByNumber(6) ; {F1} --  SwitchToDesktop num
    7::switchDesktopByNumber(7) ; {F1} --  SwitchToDesktop num
    8::switchDesktopByNumber(8) ; {F1} --  SwitchToDesktop num
    9::switchDesktopByNumber(9) ; {F1} --  SwitchToDesktop num
;}

;{ Clipboard trimming
    ^t::  ; {F1} -- Trim Clipboard
      len = 195
      ;~ InputBox, len, Trim to length of:,,,,100,xpos,ypos
      ;~ clipboard := TrimStringToSize(len,clipboard)
      clipboard = %Clipboard%
      clipboard := SubStr(clipboard, 1, len)
      DisplayTooltip(clipboard,3000,1)
      return
    t::  ; {F1} -- Trim text in current field.
      Send, ^a
      Send, ^c
      len = 195
      ;~ InputBox, len, Trim to length of:,,,,100,xpos,ypos
      ;~ clipboard := TrimStringToSize(len,clipboard)
      temp = %Clipboard%
      clipboard = %Clipboard%
      clipboard := SubStr(clipboard, 1, len)
      DisplayTooltip("Trimmed Current Field to 195 chars",3000,1)
      Send, ^v
      Clipboard = %temp%
      return
;}

;   Custom Clipboard
;{
    q:: ; {F1} -- Custom Clipboard - Sets
       q_ClipB = %clipboard%
       flash_Screen("Blue")
       DisplayTooltip(q_ClipB,1500,1)
       return
    w:: ; {F1} -- Custom Clipboard - Sets
       w_ClipB = %clipboard%
       flash_Screen("Blue")
       DisplayTooltip(w_ClipB,1500,1)
       return
    e:: ; {F1} -- Custom Clipboard - Sets
       e_ClipB = %clipboard%
       flash_Screen("Blue")
       DisplayTooltip(e_ClipB,1500,1)
       return
    a:: ; {F1} -- Custom Clipboard - Sets
       a_ClipB = %clipboard%
       flash_Screen("Blue")
       DisplayTooltip(a_ClipB,1500,1)
       return
    s:: ; {F1} -- Custom Clipboard - Sets
       s_ClipB = %clipboard%
       flash_Screen("Blue")
       DisplayTooltip(s_ClipB,1500,1)
       return
    d:: ; {F1} -- Custom Clipboard - Sets
       d_ClipB = %clipboard%
       flash_Screen("Blue")
       DisplayTooltip(d_ClipB,1500,1)
       return
    ^q:: ; {F1} -- Custom Clipboard - Recalls into Windows Clipboard
       clipboard = %q_ClipB%
       flash_Screen("Green")
       DisplayTooltip(clipboard,1500,1)
       return
    ^w:: ; {F1} -- Custom Clipboard - Recalls into Windows Clipboard
       clipboard = %w_ClipB%
       flash_Screen("Green")
       DisplayTooltip(clipboard,1500,1)
       return
    ^e:: ; {F1} -- Custom Clipboard - Recalls into Windows Clipboard
       clipboard = %e_ClipB%
       flash_Screen("Green")
       DisplayTooltip(clipboard,1500,1)
       return
    ^a:: ; {F1} -- Custom Clipboard - Recalls into Windows Clipboard
       clipboard = %a_ClipB%
       flash_Screen("Green")
       DisplayTooltip(clipboard,1500,1)
       return
    ^s:: ; {F1} -- Custom Clipboard - Recalls into Windows Clipboard
       clipboard = %s_ClipB%
       flash_Screen("Green")
       DisplayTooltip(clipboard,1500,1)
       return
    ^d:: ; {F1} -- Custom Clipboard - Recalls into Windows Clipboard
       clipboard = %d_ClipB%
       flash_Screen("Green")
       DisplayTooltip(clipboard,1500,1)
       return
;}
    !n::createVirtualDesktop()  ;  {F1} --  New Virtual Desktop
;    m::deleteVirtualDesktop()
}
#If

F2::  ; zzz_ Control Character for shortcuts - VSA / Liveconnect --- Note: this key triggers VSAExportGetFileHeaders()
{
 VSA_Indexes := VSAExportGetFileHeaders()  ; Needed for VSA Agent Export lookups to function
 return
 #If (A_PriorHotKey = "F2" AND A_TimeSincePriorHotkey < User_Config_File.Timer.FN_Key_Timer)
}
#If

F3:: return ; zzz_ Control Character for shortcuts - Credential Manager / Password Management System Passord Lookups
{
#If (A_PriorHotKey = "F3" AND A_TimeSincePriorHotkey < User_Config_File.Timer.FN_Key_Timer)

    3:: ; {F3} -- Password - Random Password Generator
    ;{
      f3a_ClipB := GetRandPassword(25)
      Clipboard := f3a_ClipB
      DisplayTooltip(f3a_ClipB,1000,1)
      return
    ;}
    -:: ; {F3} -- Credential_Manager - Copies user's domain credential
    ;{
       Data = ;nul
       Data := CredRead("MY_COMPANY_NAME_AD_cred")
       If (Data.name = "")
       {
         flash_Screen("Red")
       } else {
         flash_Screen("Blue")
       }
       return
    ;}
    e::   ; {F3} -- Credential_Manager - Copies user's Office 365 creds  e = email
    ;{
       Data = ;nul
       Data := CredRead("Office365_cred")
       If (Data.name = "")
       {
         flash_Screen("Red")
       } else {
         flash_Screen("Blue")
       }
       return
    ;}
    g::    ; {F3} -- GENERIC PASSWORD RESET  Raining123!!@
    ;{
      clipboard := User_Config_File.Win_Cred_Man.GenericPasswordResets
      DisplayTooltip(clipboard,User_Config_File.Timer.PasswordCopied_Tooltip,1)
      return
    ;}
    s::   ; {F3} -- Credential_Manager - company123 domain admin creds
    ;{
       Data = ;nul
       Data := CredRead("company123_Admin_cred")
       If (Data.name = "")
       {
         flash_Screen("Red")
       } else {
         clipboard := Data.username . "`t" . Data.password
         flash_Screen("Blue")
         DisplayTooltip("company123 Admin Username and Password Copied",User_Config_File.Timer.UsernameCopied_Tooltip,1)
       }
       return
    ;}
    p::   ; {F3} -- Credential_Manager - EmployeeTimeTrackingSoftware creds
    ;{
       Data = ;nul
       Data := CredRead("EmployeeTimeTrackingSoftware_cred")
       If (Data.name = "")
       {
         flash_Screen("Red")
       } else {
         flash_Screen("Blue")
       }
       return
    ;}
    i::  ; {F3} -- Password Management System --- Opens Password Management System for organization
    ;{
      MouseGetPos,xpos,ypos,cursWin,cursControl
      WinGetTitle, UnderMouse_WindowTitle, ahk_id %cursWin%
      If ( InStr(UnderMouse_WindowTitle, "::Shared") || InStr(UnderMouse_WindowTitle, "::Private") )
      {
        StringSplit, part, UnderMouse_WindowTitle, :
        FullGroupName = %part1%
        RegExMatch(FullGroupName, ".*\.(.*)", r)
        VSA_GroupName = %r1%
        Password Management System_Organization := Organization_Lookup_Function(VSA_GroupName)
        Password Management System_Organization := Password Management System_Organization.Password Management System_Org_ID
        DisplayTooltip(VSA_GroupName,1000,1)
      } else {
        Password Management System_Organization := Organization_Lookup_Function(clipboard)
        Password Management System_Organization := Password Management System_Organization.Password Management System_Org_ID
      }
      If (Password Management System_Organization = "")
        {
          flash_Screen("Red")
          Run, %Kaseya_Browser% "https://MY_COMPANY_NAME.Password Management Systemlue.com/organizations"
        } else {
          Run, %Kaseya_Browser% "https://MY_COMPANY_NAME.Password Management Systemlue.com/%Password Management System_Organization%"
      }
      return
    ;}
    f::  ; {F3} -- "File Sharing"  Password Management System --- Opens File Sharing in Password Management System for organization
    ;{
      MouseGetPos,xpos,ypos,cursWin,cursControl
      WinGetTitle, UnderMouse_WindowTitle, ahk_id %cursWin%
      If ( InStr(UnderMouse_WindowTitle, "::Shared") || InStr(UnderMouse_WindowTitle, "::Private") )
      {
        StringSplit, part, UnderMouse_WindowTitle, :
        FullGroupName = %part1%
        RegExMatch(FullGroupName, ".*\.(.*)", r)
        VSA_GroupName = %r1%
        Password Management System_Organization := Organization_Lookup_Function(VSA_GroupName)
        Password Management System_Organization := Password Management System_Organization.Password Management System_Org_ID
        DisplayTooltip(VSA_GroupName,1000,1)
      } else {
        Password Management System_Organization := Organization_Lookup_Function(clipboard)
        Password Management System_Organization := Password Management System_Organization.Password Management System_Org_ID
        if (Password Management System_Organization = "")
        {
          if WinActive("ahk_exe chrome.exe") or WinActive("ahk_exe msedge.exe")
          {
            temp = %clipboard%
            Send, ^a
            Send, ^c
            Password Management System_Organization := Organization_Lookup_Function(clipboard)
            Password Management System_Organization := Password Management System_Organization.Password Management System_Org_ID
            clipboard = %temp%
          }
        }
      }
      If (Password Management System_Organization = "")
      {
        DisplayTooltip("Organization not found...`n`nHover Over VSA Remote Controll Session`nor`nCopy Full Company Name",3000,1)
      } else {
        Run, %Kaseya_Browser% "https://MY_COMPANY_NAME.Password Management Systemlue.com/%Password Management System_Organization%/assets/269145-file-sharing/records"
      }
      return
    ;}
    o::   ; {F3} --  O365 creds - Copy FULL company name from Password Management System or TicketingSystem to clipboard prior to hotkey
    ;{
      MouseGetPos,xpos,ypos,cursWin,cursControl
      WinGetTitle, UnderMouse_WindowTitle, ahk_id %cursWin%
      If ( InStr(UnderMouse_WindowTitle, "::Shared") || InStr(UnderMouse_WindowTitle, "::Private") )
      {
        StringSplit, part, UnderMouse_WindowTitle, :
        FullGroupName = %part1%
        RegExMatch(FullGroupName, ".*\.(.*)", r)
        VSA_GroupName = %r1%
        DisplayTooltip(VSA_GroupName,3000,1)
        Password Management System_Credential := Organization_Lookup_Function(VSA_GroupName)
        Password Management System_Credential := Password Management System_Credential.O365_Cred
      } else {
        Password Management System_Credential := Organization_Lookup_Function(clipboard)
        Password Management System_Credential := Password Management System_Credential.O365_Cred
      }
      Run, %Office_Browser% -incognito "portal.office.com"
      if (Password Management System_Credential) {
        StringSplit, part, Password Management System_Credential, _
        x_URL=https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%
        if (User_Config_File.Rufaydium.CtrlAltMiddleButton_UseRufaydium = 0) {
          StringSplit, part, Password Management System_Credential, _
           Run %Kaseya_Browser% %x_URL%
        } else {
          Data := Rufaydium_Credential_Lookup(Password Management System_Credential, User_Config_File.Rufaydium.CtrlAltMiddleButton_HeadlessMode, User_Config_File.Rufaydium.O365_IsUserAuthenticated_Sleep, User_Config_File.Rufaydium.Password Management System_Default_SleepTimer, User_Config_File.Timer.UsernameCopied_Tooltip, User_Config_File.Rufaydium.CtrlAltMiddleButton_CloseAfterReadAttempt)
          TOTP := Data.TOTP
          if (Data.name != "") {
            Clipboard := Data.username
            DisplayTooltip(clipboard,User_Config_File.Timer.UsernameCopied_Tooltip,1)
          }
        }
      } else {
        DisplayTooltip("No Password Management System credential found",User_Config_File.Timer.UsernameCopied_Tooltip,1)
      }
      return
    ;}
    /::  ; {F3} --  Writes cred to Cred Manager.  Chrome, on the Password Management System URL. have username copied to F1 -> q   &   have password copied to  F1 -> a
    ;{
      URL := getChromeTabURL()
      If InStr(URL, "MY_COMPANY_NAME.Password Management Systemlue.com")
      {
        pos := RegExMatch(URL, "(\d+.*/\d+)", m)
        StringSplit, part, m, /
        Password Management System_Credential = %part1%_%part3%
        CredWrite(Password Management System_Credential,q_ClipB,a_ClipB)
        flash_Screen("Blue")
      } else {
        flash_Screen("Red")
      }
      return
    ;}
    ,::  ; {F3} --  DEVELOPER USE -- (Chrome) If on Password Management System URL - See if Cred Exists in Windows Cred Manager --  Copies %Password Management System_Credential% to clipboard
    ;{
      URL := getChromeTabURL()
      If InStr(URL, "MY_COMPANY_NAME.Password Management Systemlue.com")
      {
        pos := RegExMatch(URL, "(\d+.*/\d+)", m)
        StringSplit, part, m, /
        Password Management System_Credential = %part1%_%part3%
        clipboard = %Password Management System_Credential%
        ;~ clipboard = %part1%
        DisplayTooltip(clipboard,1500,1)
        Data := CredRead(Password Management System_Credential)
        If (Data.name = "")
        {
          flash_Screen("Red")
          flash_Screen("Red")
          flash_Screen("Red")
        } else {
          flash_Screen("Green")
        }
        return
      }
      return
    ;}
}
#If

F4::  return  ; zzz_ Control Character for shortcuts - Rufaydium Web-Driver
{
#If (A_PriorHotKey = "F4" AND A_TimeSincePriorHotkey < User_Config_File.Timer.FN_Key_Timer)
    /:: ; {F4} --  Initial Credential Manager Setup --- Resync (NOCAdmin,O365,RMM) credentials that the script references
;{
      MSEdge := new Rufaydium("msedgedriver.exe","--port=54226")
      MSEdge.capabilities.HeadlessMode := User_Config_File.Rufaydium.HeadlessMode
      ;~ MSEdge.Maximize()
      Cred_Man_Init_Session := MSEdge.NewSession()
      Cred_Man_Init_Session.Navigate("https://MY_COMPANY_NAME.Password Management Systemlue.com")
      Sleep User_Config_File.Rufaydium.O365_IsUserAuthenticated_Sleep
      If InStr(Cred_Man_Init_Session.url, "login.microsoftonline.com") {
        Data = ;nul
        Data := CredRead("Office365_cred")
        If (Data.name = "")
        {
          flash_Screen("Red")
        } else {
          f3q_ClipB := Data.username
          f3a_ClipB := Data.password
          clipboard = %f3q_ClipB%
          flash_Screen("Blue")
          DisplayTooltip("Personal Credential Read`n" . Office365_cred,User_Config_File.Timer.UsernameCopied_Tooltip,1)
        }
        MsgBox, Please Login To Office Before Clickng Ok
      }
      CredList_URLs=  ; nul
      Accounts = ; nul
      Loop, read, C:\Users\%username%\Desktop\Scripts\Reference_Files\Script_Organization_Lookup.tsv
      {
        LineNumber := A_Index
        if (LineNumber = 1) {     ; Get indexes
          Loop, parse, A_LoopReadLine, %A_Tab%
          {
            If InStr(A_LoopField, "Account") {
              Account_Index = %A_Index%
            }
          }
        } else {
          Loop, parse, A_LoopReadLine, %A_Tab%
          {
            If (A_Index = Account_Index) {
              Accounts .= A_LoopField . "`n"
            }
          }
        }
      }
      ;~ msgbox, %Accounts%
      Loop, parse, Accounts, `n, `r
      {
        DACred=
        ADACred=
        O365Cred=
        temp := Organization_Lookup_Function(A_LoopField)
        DACred := temp.NOC_Cred
        ADACred := temp.RMM_Cred
        O365Cred := temp.O365_Cred
        if (DACred) {
          StringSplit, part, DACred, _
          URL = https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%`n
          CredList_URLs .= URL
        }
        if (ADACred) {
          StringSplit, part, ADACred, _
          URL = https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%`n
          CredList_URLs .= URL
        }
        if (O365Cred) {
          StringSplit, part, O365Cred, _
          URL = https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%`n
          CredList_URLs .= URL
        }
      }
      ;~ FileAppend, %CredList_URLs%, C:\Users\%username%\Desktop\Scripts\temp_temp.txt
      ;~ Run, C:\Users\%username%\Desktop\Scripts\temp_temp.txt
      CredList_URLs .= "https://MY_COMPANY_NAME.Password Management Systemlue.com/7164431/passwords/22778376`n"  ; RestorePoint Teams
      CredList_URLs .= "https://MY_COMPANY_NAME.Password Management Systemlue.com/7164431/passwords/22778362`n"  ; RestorePoint Manage Engine
      Sort CredList_URLs, U D`n   ; Sort - U removes duplicates D delimeter newline
      StringReplace CredList_URLs,CredList_URLs,`n,`n,UseErrorLevel
      total_URLS = %ErrorLevel%
      total_URLS += 1
      count = 0
      Loop, parse, CredList_URLs, `n, `r
      {
        count += 1
        Cred_Man_Init_Session.Navigate(A_LoopField)
        sleep User_Config_File.Rufaydium.Password Management System_Default_SleepTimer
        if (A_Index = 1) {
          sleep User_Config_File.Rufaydium.Password Management System_First_URL_Load_SleepTimer  ; Extra time for first URL
        }
        User := Cred_Man_Init_Session.querySelector("#react-main > div > div:nth-child(3) > div > div.col-sm-8 > div > dl:nth-child(2) > div.row-with-copy.qa-username-row > div.row-content.qa-row-content > span").innerText ; Copy Username
        ;~ MsgBox % User
        ;~ MsgBox, %User%
        Cred_Man_Init_Session.querySelector("#react-main > div > div:nth-child(3) > div > div.col-sm-8 > div > dl:nth-child(2) > div.show-password-in-vault > div.show-password-row.qa-show-password-row.dl-inner-container > dd > div.show-password-wrapper > div > span").click()   ; Click "Show Password"
        Sleep 1000
        Pass := Cred_Man_Init_Session.querySelector("#react-main > div > div:nth-child(3) > div > div.col-sm-8 > div > dl:nth-child(2) > div.show-password-in-vault > div.show-password-row.qa-show-password-row.dl-inner-container > dd > div.password-revealed > input").value
        ;~ MsgBox % User . "`n" . Pass . "`n" . Cred_Man_Init_Session.url
        If InStr(Cred_Man_Init_Session.url, "MY_COMPANY_NAME.Password Management Systemlue.com")
        {
          pos := RegExMatch(Cred_Man_Init_Session.url, "(\d+.*/\d+)", m)
          StringSplit, part, m, /
          Password Management System_Credential_Automation = %part1%_%part3%
          ;~ MsgBox % m . "`n" . Password Management System_Credential_Automation . "`n" . User . "`n" . Pass . "`n" . Cred_Man_Init_Session.url
          units=seconds
          ETR := 5 * (total_URLS - count)
          if (ETR > 60) {
            ETR := ETR / 60
            StringSplit, part, ETR, .
            part2 := SubStr(part2, 1, 1)
            ETR := part1 . "." . part2
            units=minutes
          }
          if ( (User) && (Pass) ) {
            CredWrite(Password Management System_Credential_Automation,User,Pass)
            A := Mod(count, 4)
            if ( (A = 0) || (units = "seconds") ){   ; Rate Limit Notifications - Every 4 - Unless script is close to done
              DisplayTooltip("Credential Wrote - " . count . " / " . total_URLS . "`nETR --- " . ETR . " " . units,4000,1)
            }
          } else {
            DisplayTooltip("Unable to Find Username and Password - " . Password Management System_Credential_Automation,4000,1)
          }
        }
      }
      Loop, Files, C:\Users\%username%\AppData\Local\Temp\*.*, D
      {
        if InStr(A_LoopFileFullPath, "scoped_dir") || InStr(A_LoopFileFullPath, "edge_BITS_") {   ; Remove Edge Specific Temp Files
          ;~ DisplayTooltip(A_LoopFileFullPath,1000,1)
          FileRemoveDir, % A_LoopFileFullPath, 1
        }
      }
      flash_Screen_Rainbow()
    return
;}
    !/:: ; {F4} --  Resync ALL credentials that the script can possibly reference
;{
;{  Initial Connection  and  Authentication to Office Portal
      MSEdge := new Rufaydium("msedgedriver.exe","--port=54226")
      MSEdge.capabilities.HeadlessMode := User_Config_File.Rufaydium.HeadlessMode
      ;~ MSEdge.Maximize()
      Cred_Man_Init_Session := MSEdge.NewSession()
      Cred_Man_Init_Session.Navigate("https://MY_COMPANY_NAME.Password Management Systemlue.com")
      Sleep User_Config_File.Rufaydium.O365_IsUserAuthenticated_Sleep
      If InStr(Cred_Man_Init_Session.url, "login.microsoftonline.com") {
        Data = ;nul
        Data := CredRead("Office365_cred")
        If (Data.name = "")
        {
          flash_Screen("Red")
        } else {
          f3q_ClipB := Data.username
          f3a_ClipB := Data.password
          clipboard = %f3q_ClipB%
          flash_Screen("Blue")
          DisplayTooltip("Personal Credential Read`n" . Office365_cred,User_Config_File.Timer.UsernameCopied_Tooltip,1)
        }
        MsgBox, Please Login To Office Before Clickng Ok
      }
;}
;{  Preliminary Check - Ensure Columns Exists --- Name, Type , StatusGet Header Row indexes on  https://MY_COMPANY_NAME.Password Management Systemlue.com/organizations
      Cred_Man_Init_Session.Navigate("https://MY_COMPANY_NAME.Password Management Systemlue.com/organizations")
      sleep 1500  ; Load time
      HeaderRow := Cred_Man_Init_Session.querySelector(".react-table-head.qa-react-table-head > div > table > thead > tr")
      FullClientTable := Cred_Man_Init_Session.querySelector(".react-table-body.qa-react-table-body.position-relative > div > table > tbody")
      temp := HeaderRow.InnerHTML
      ;~ msgbox, %temp%
      StringReplace, temp,temp,</th>,</th>`n, UseErrorLevel    ; Add New Lines For Readability (Replace All Occurences)
      FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Active_Client_Listing.txt
      FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Client_Table_Headers_InnerHTML.txt
      FileAppend, %temp%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Client_Table_Headers_InnerHTML.txt
      Name_I= ;nul
      Type_I= ;nul
      Status_I= ;nul
      ;~ msgbox, %temp%
      Loop, Read, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Client_Table_Headers_InnerHTML.txt
      {
        ;~ MsgBox, %A_Index%`n`n%A_LoopReadLine%
        If ( InStr(A_LoopReadLine, "column-name asc") ) {
          ;~ MsgBox, %A_Index%`n`nName Column
          Name_I=%A_Index%
        } else if ( InStr(A_LoopReadLine, "column-organization-type-name") ) {
          Type_I=%A_Index%
          ;~ MsgBox, %A_Index%`n`nType Column
        } else if ( InStr(A_LoopReadLine, "column-organization-status-name") ) {
          Status_I=%A_Index%
          ;~ MsgBox, %A_Index%`n`nStatus Column
        }
      }
      if ( (Name_I = "") || (Type_I = "") || (Status_I = "") ) {
        MsgBox, Error: The following rows must exist in Password Management System's Organization view/filter`n`nName`nType`nStatus`n`nCurrent Thread Exiting.
        return
      }
;}
;{  Check each row to see if Active Client
      temp := FullClientTable.InnerHTML
      ;~ msgbox, %temp%
      StringReplace, temp,temp,</tr><tr,</tr>`n<tr, UseErrorLevel    ; Split Table Rows - Add New Lines For Readability (Replace All Occurences)
      FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Client_Table_Listing_InnerHTML.txt
      FileAppend, %temp%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Client_Table_Listing_InnerHTML.txt
      Loop, Read, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Client_Table_Listing_InnerHTML.txt
      {
        URL= ;nul
        org= ;nul
        line= ;nul
        client= ;nul
        active= ;nul
        if (RegExMatch(A_LoopReadLine, ".*a\s+href=(.*)<\/a><\/td>", SubPat)) {  ; Organization URL
          ;~ MsgBox, Regex Matched`n`n%SubPat1%
          StringReplace, SubPat1,SubPat1,", , All    ; Remove Quotes from match (All Occurences)   ; Example  "/6536407">Boehm Pressed Steel Company
          StringSplit, part, SubPat1, >
          URL=%part1%
          org=%part2%
          StringReplace,org,org,amp;,, All   ; Make Ampersand appear normally
          line=%org%`t%URL%`n
          if (RegExMatch(A_LoopReadLine, ".*\s+title=.Client.><span>(.*)<\/span>", client)) {  ; Is Client?
            if (RegExMatch(A_LoopReadLine, ".*\s+title=.Active.><span>(.*)<\/span>", active)) {  ; Is Active?
              FileAppend, %line%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Active_Client_Listing.txt
            } else {
              ;~ MsgBox % "URL = " . URL . "`norg = " . org . "`n`nCompany is not active"
            }
          } else { ; Current Company in not a Client
            ;~ MsgBox % "URL = " . URL . "`norg = " . org . "`n`nCompany is not a client"
            continue
          }
        }
      }
;}
;{  Initialize Files in preperation to gather useful data
      FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv
      FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_successful_nslookups.txt
      FileDelete,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_failed_pings.txt
      FileDelete,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_successful_pings.txt
      FileDelete,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_failed_nslookups.txt
      FileDelete,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_pingable_URL_list.txt
      FileDelete,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_href_list.txt
      line=URL`thref`tAccount`tCredential`tName`n
      FileAppend,%line%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv
;}
      Loop, Read, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Active_Client_Listing.txt    ; Loop over All Active Clients
      {
;{  Navigate to client's Password Management System "Passwords" Page
        StringSplit, part, A_LoopReadLine, `t
        org=%part1%
        URL := "https://MY_COMPANY_NAME.Password Management Systemlue.com" . part2 . "/passwords"
        Cred_Man_Init_Session.Navigate(URL)
        DisplayTooltip(org . "`nCollecting a list of All Passwords with valid URLS",3000,1)
        sleep User_Config_File.Rufaydium.Password Management System_Default_SleepTimer
;}
;{  If First Client -- Ensure "URL" exists or gets added to Filter
        if (A_Index = 1) {   ; Special Check  --- Need to make sure the URL is listed in the Tech's view/filter
          URL_I= ;nul
          HeaderRow := Cred_Man_Init_Session.querySelector(".react-table-head.qa-react-table-head > div > table > thead > tr")
          temp := HeaderRow.InnerHTML
          StringReplace, temp,temp,</th>,</th>`n, UseErrorLevel    ; Add New Lines For Readability (Replace All Occurences)
          FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Password_Table_Headers_InnerHTML.txt
          FileAppend, %temp%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Password_Table_Headers_InnerHTML.txt
          ;~ Run, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Password_Table_Headers_InnerHTML.txt
          i=0
          Loop, Read, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Organization_Password_Table_Headers_InnerHTML.txt
          {
            if (RegExMatch(A_LoopReadLine, "\s+title=.(.*).\s+style=", SubPat)) {
              i += 1
              ;~ MsgBox, Header Row --- Looking for URL`n%i% -- %SubPat1%
              If InStr(SubPat1, "URL") {
                URL_I := i
                ;~ MsgBox % "URL_I = " . URL_I
              }
            }
          }
          if (URL_I = "") {
            ;~ MsgBox, URL doesn't exist upon loading - Trying to add URL
            Cred_Man_Init_Session.querySelector(".table-with-filter-actions.display-flex > button > i").click()
            sleep 300
            AllCheckboxes := Cred_Man_Init_Session.querySelector(".dialog-body.qa-dialog-body > div > div > div:nth-child(2)")
            temp := AllCheckboxes.InnerHTML
            StringReplace, temp,temp,</div></div></div><div,</div></div></div>`n<div, UseErrorLevel    ; Split elements - Add New Lines For Readability (Replace All Occurences)
            FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Password_ListViewOptions.txt
            FileAppend, %temp%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Password_ListViewOptions.txt
            i=0
            Loop, Read, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Password_ListViewOptions.txt
            {
              if (RegExMatch(A_LoopReadLine, "span\s+class=.field-label.>(.*)<\/span>", SubPat)) {
                i += 1
                ;~ MsgBox, %i% -- %SubPat1%
                If InStr(SubPat1, "URL") {
                  ;~ MsgBox, URL Found - %i%
                  URL_I := i
                  Cred_Man_Init_Session.querySelector(".dialog-body.qa-dialog-body > div > div > div:nth-child(2) > div:nth-child(" . i . ") > div > div > label > div > div > input").click()
                  Cred_Man_Init_Session.querySelector(".react-button.qa-react-button.default.apply-button.qa-apply-button").click()
                }
              }
            }
          }
          if (URL_I = "") {
            MsgBox, Error: The URL column must exist in Password Management System's Passwords page view/filter`n`nCurrent Thread Exiting.
            return
          }
        }
;}
;{  Include functionality to write all NOC Admin - NOC RMM - and O365 creds for all Clients
        DACred=
        ADACred=
        O365Cred=
        temp := Organization_Lookup_Function(A_LoopField)
        DACred := temp.NOC_Cred
        ADACred := temp.RMM_Cred
        O365Cred := temp.O365_Cred
        if (DACred) {
          StringSplit, part, DACred, _
          URL = https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%`n
          FileAppend,URL,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_href_list.txt
        }
        if (ADACred) {
          StringSplit, part, ADACred, _
          URL = https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%`n
          FileAppend,URL,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_href_list.txt
        }
        if (O365Cred) {
          StringSplit, part, O365Cred, _
          URL = https://MY_COMPANY_NAME.Password Management Systemlue.com/%part1%/passwords/%part2%`n
          FileAppend,URL,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_href_list.txt
        }
        ;~ MsgBox % DACred . "`n" . ADACred . "`n" . O365Cred
;}
;{  Gather a List of ALL passwords that contain a "URL" that can be accessed from a computer at MY_COMPANY_NAME.  Stores in C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv
    ;~ Excludes Private IP addresses for all client companies except MY_COMPANY_NAME  as a tech cannot access other companies private resourses from their PC
        PasswordTable := Cred_Man_Init_Session.querySelector(".react-table-body.qa-react-table-body.position-relative > div > table > tbody")
        temp := PasswordTable.InnerHTML
        StringReplace, temp,temp,</tr><tr data,</tr>`n<tr data, UseErrorLevel    ; Split elements - Add New Lines For Readability (Replace All Occurences)
        FileDelete, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Passwords_Listed.txt
        FileAppend, %temp%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Passwords_Listed.txt
        ;~ Run, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Passwords_Listed.txt
        ;~ msgbox, %temp%
        ;~ i=0
        Loop, Read, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_Passwords_Listed.txt
        {
          URL= ;nul
          if (RegExMatch(A_LoopReadLine, "<a\s+class=.display-inline.\s+target=._new.\s+href=(.*)<\/a><\/td>", SubPat)) {
            ;~ i += 1
            ;~ MsgBox, %SubPat1%
            StringSplit, part, SubPat1, >
            StringReplace, part1,part1,", , All    ; Remove double Quotes
            URL := part1
            if InStr(URL, "http://null") {  ; No URL Defined on Password Management System asset
              continue
            } else {  ; URL is defined on current row  --- Use Regex to extract data     variables: URL - href - cred - Name
              if RegExMatch(URL, "(127\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3})|(10\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3})|(172\.1[6-9]{1}[0-9]{0,1}\.[0-9]{1,3}\.[0-9]{1,3})|(172\.2[0-9]{1}[0-9]{0,1}\.[0-9]{1,3}\.[0-9]{1,3})|(172\.3[0-1]{1}[0-9]{0,1}\.[0-9]{1,3}\.[0-9]{1,3})|(192\.168\.[0-9]{1,3}\.[0-9]{1,3})" ) {
              ; URL points to an Internal IP Address
                if !InStr(org, "MY_COMPANY_NAME , LLC") {  ; If company is not MY_COMPANY_NAME
                  continue  ; We cant access other companies internal IPs
                }
              } else {
                if ( (InStr(org, "Riten Industries, Inc.")) && (InStr(URL, "192.0.0.")) ) {   ; Riten uses non standard private IP addressing... Somebody should change this.
                  continue
                }
              }
              StringSplit, z, URL, /
              StringSplit, part, z3, :
              pingable_URL := part1
              if (RegExMatch(A_LoopReadLine, "class=.actions-cell-content.><a\shref=.(.*)\/edit.><i\s+class=.row-action-icon", SubPat)) {
                ;~ MsgBox % SubPat1
                href := "https://MY_COMPANY_NAME.Password Management Systemlue.com" . SubPat1
                StringSplit, part, SubPat1, /
                cred = %part2%_%part4%
              }
              if (RegExMatch(A_LoopReadLine, "ellipsis\s+name-link.>(.*)<\/a><\/div", SubPat)) {
                Name := SubPat1
              }
;{  ping / nslookup  tests
              line=%URL%`t%href%`t%org%`t%cred%`t%Name%`n
              StringReplace, line,line,amp;, , All
              if (RegExMatch(pingable_URL, "(\b25[0-5]|\b2[0-4][0-9]|\b[01]?[0-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}")) {   ; URL is an IP Address
                RunWait cmd.exe /c ping %pingable_URL% -n 2 1> C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_ping.txt 2>&1, , hide
                FileRead, temp, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_ping.txt
                if (RegExMatch(temp, "Reply from (\b25[0-5]|\b2[0-4][0-9]|\b[01]?[0-9][0-9]?)(\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}: bytes=\d+ time=\d+ms TTL=\d+", z)) {   ; Checks if Ping was successful
                  FileAppend,%line%`n%temp%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_successful_pings.txt
                } else {
                  FileAppend,%line%`n%temp%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_failed_pings.txt
                  ;~ FileAppend,%line%`n%temp%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_ping.txt
                  ;~ Run, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_ping.txt
                  continue
                }
              }
              RunWait cmd.exe /c nslookup %pingable_URL% 1> C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_nslookup.txt 2>&1, , hide   ; nslookup hostname - Redirect STDERR to STDOUT
              FileRead, temp, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_nslookup.txt
              if (RegExMatch(temp, "s)Name:\s+(?<protocol>[a-z\-]+)?:?(?<slash>\/\/)?(?<host>[-a-zA-Z0-9\.]+)?(?P<port>:\d+)?\s*.Address(es)?:", z)) {   ; Checks if DNS lookup was successful
                ;~ MsgBox % "DNS Lookup Successful`n`n" . temp
                FileAppend,%line%`n%temp%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_successful_nslookups.txt
              } else {
                FileAppend,%line%`n%temp%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_failed_nslookups.txt
                ;~ FileAppend,`n%line%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_nslookup.txt
                ;~ Run, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_nslookup.txt
                ;~ MsgBox % "DNS Lookup Failed`n`n" . Name . "`n`n" . temp     ; Not writing Cred since DNS lookup failed
                continue
              }
              y_URL := pingable_URL
              if (RegExMatch(y_URL,"^\.") ) {
              } else {
                x_URL= ;nul
                lf= ;nul
                cf= ;nul
                array := StrSplit(y_URL, ".")
                while !(RegExMatch(y_URL,"^\.")) {
                  if (A_Index > 4) {
                    break
                  }
                  cf := array[A_Index]
                  if (A_Index = 1) {
                    lf := cf
                  } else {
                    if ( (cf = "com") || (cf = "org") || (cf = "net") || (cf = "io") || (cf = "us") || (cf = "cloud")  ) {
                      y_URL := "." . lf . "."
                    }
                    lf := cf
                    cf := array[A_Index]
                    if !(cf) {
                      break
                    }
                  }
                }
              }
;}
;{  Write successful credential
              ;~ MsgBox % line
              FileAppend,%y_URL%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_y_URL_list.txt    ; This is what  Valid_URL_Credential_Lookup  should condense the URL and perform a lookup on
              FileAppend,%pingable_URL%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_pingable_URL_list.txt
              FileAppend,%href%`n,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_href_list.txt
              FileAppend,%line%,C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv
              ;~ MsgBox % "URL = " . URL . "`nhref = " . href
;}
            }
          }
        }
;}
      }
;{  Copy Final Results (TSV File) to M: and C: drive folders
      if (User_Config_File.Maintain_Script.Are_You_Wanting_To_Maintain_The_Scripts_Reference_Files) {
        Run cmd.exe /c copy "C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv" "M:\Shared_SMB\Folder_Path\Reference_Files\Valid_Client_URL_Cred_Lookup.tsv" /y /d, , hide
        Run cmd.exe /c copy "C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv" "C:\Users\%username%\Desktop\Scripts\Reference_Files\Valid_Client_URL_Cred_Lookup.tsv" /y /d, , hide
      }
;}
;{  Loop Over a list of all valid URLs
      FileRead, CredList_URLs, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_Password Management System_href_list.txt
      Sort CredList_URLs, U D`n   ; Sort - U removes duplicates D delimeter newline
      StringReplace CredList_URLs,CredList_URLs,`n,`n,UseErrorLevel
      total_URLS = %ErrorLevel%
      total_URLS += 1
      count = 0
      Loop, parse, CredList_URLs, `n, `r
      {
        count += 1
        Cred_Man_Init_Session.Navigate(A_LoopField)
        sleep User_Config_File.Rufaydium.Password Management System_Default_SleepTimer
        if (A_Index = 1) {
          sleep User_Config_File.Rufaydium.Password Management System_First_URL_Load_SleepTimer  ; Extra time for first URL
        }
        User := Cred_Man_Init_Session.querySelector(".row-with-copy.qa-username-row > div.row-content.qa-row-content > span").innerText ; Copy Username
        ;~ MsgBox % User
        ;~ MsgBox, %User%
        Cred_Man_Init_Session.querySelector(".show-password-in-vault > div.show-password-row.qa-show-password-row.dl-inner-container > dd > div.show-password-wrapper > div > span").click()   ; Click "Show Password"
        Sleep 1000
        Pass := Cred_Man_Init_Session.querySelector(".show-password-in-vault > div.show-password-row.qa-show-password-row.dl-inner-container > dd > div.password-revealed > input").value
        ;~ MsgBox % User . "`n" . Pass . "`n" . Cred_Man_Init_Session.url
        If InStr(Cred_Man_Init_Session.url, "MY_COMPANY_NAME.Password Management Systemlue.com")
        {
          pos := RegExMatch(Cred_Man_Init_Session.url, "(\d+.*/\d+)", m)
          StringSplit, part, m, /
          Password Management System_Credential_Automation = %part1%_%part3%
          ;~ MsgBox % m . "`n" . Password Management System_Credential_Automation . "`n" . User . "`n" . Pass . "`n" . Cred_Man_Init_Session.url
          units=seconds
          ETR := 5 * (total_URLS - count)
          if (ETR > 60) {
            ETR := ETR / 60
            StringSplit, part, ETR, .
            part2 := SubStr(part2, 1, 1)
            ETR := part1 . "." . part2
            units=minutes
          }
          if ( (User) && (Pass) ) {
            CredWrite(Password Management System_Credential_Automation,User,Pass)
            A := Mod(count, 4)
            if ( (A = 0) || (units = "seconds") ){   ; Rate Limit Notifications - Every 4 - Unless script is close to done
              DisplayTooltip("Credential Wrote - " . count . " / " . total_URLS . "`nETR --- " . ETR . " " . units,4000,1)
            }
          } else {
            DisplayTooltip("Unable to Find Username and Password - " . Password Management System_Credential_Automation,4000,1)
          }
        }
      }
      ;~ Run, C:\Users\%username%\Desktop\Scripts\Temp_Files\temp_FULL_PASSWORD_URL_LIST_with_HREFS_AllOrganizations.tsv


;{   Clean Up Browser Temp File
      Loop, Files, C:\Users\%username%\AppData\Local\Temp\*.*, D
      {
        DisplayTooltip("Cleaning Up Edge Temp Files",2000,1)
        if InStr(A_LoopFileFullPath, "scoped_dir") || InStr(A_LoopFileFullPath, "edge_BITS_") {   ; Remove Edge Specific Temp Files
          ;~ DisplayTooltip(A_LoopFileFullPath,1000,1)
          FileRemoveDir, % A_LoopFileFullPath, 1
        }
      }
;}
      flash_Screen_Rainbow()
    return
;}
;}
    u:: ; {F4} -- Unknown Ticket Automation -- Single Passthrough - Uses MS Edge
;{   Unknown Ticket Automation
      MSEdge_UT := new Rufaydium("msedgedriver.exe","--port=54326")
      UT_Session := MSEdge_UT.NewSession()    ; Launch Web Browser with specified capabilities
      MSEdge_UT.capabilities.HeadlessMode := User_Config_File.Rufaydium.HeadlessMode
      MSEdge_UT.Navigate("portal.office.com")
      Sleep User_Config_File.Rufaydium.O365_IsUserAuthenticated_Sleep
      If InStr(MSEdge_UT.url, "login.microsoftonline.com") {
        Data = ;nul
        Data := CredRead("Office365_cred")
        If (Data.name = "")
        {
          flash_Screen("Red")
        } else {
          f3q_ClipB := Data.username
          f3a_ClipB := Data.password
          clipboard = %f3q_ClipB%
          flash_Screen("Blue")
          DisplayTooltip("Personal Credential Read`n" . Office365_cred,User_Config_File.Timer.UsernameCopied_Tooltip,1)
        }
        MsgBox, Please Login To Office Before Clickng Ok
      }
      ;~ UT_Session.Maximize()     ; FullScreen Mode
      ;{   Authenticate to TicketingSystem
      UT_Session.Navigate("https://TicketingSystem.kaseya.com/Gateway.aspx?enc=%2fZyI3iknrKMpB8K8w6gKLbUeJsnECRcITlan3qIkV9GLYTVtQkVPNqWJY%2ftQiYkl%2fATNIRE2wODGGweAsGY%2fDQ%3d%3d")   ; TicketingSystem GateWay
      UT_Session.querySelector("#ctl00_MainContent_txtUserID").value := "cjohnson"   ; Username to login as
      UT_Session.querySelector("#ctl00_MainContent_btnNext").click()    ; Click Button to login
      ;},
      UT_Session.Navigate("https://TicketingSystem.kaseya.com/MSP/PendingTickets.aspx")    ;  Open "Unknown Tickets" (Root URL)
      All_Ticket_tbody := UT_Session.querySelector("#ctl00_phListing_ctrlListing_dgrdListing_ctl00 > tbody")  ; Unkown Ticket Listing (Parent Element)
      All_Ticket_Full_HTML := All_Ticket_tbody.InnerHTML     ; Grab the HTML from the unknown tickets
      ;~ RegExMatch(All_Ticket_Full_HTML, ".*href=(.*) data-id", SubPat)  ;
      ticket_count = 0
      FileDelete, C:\users\%username%\Documents\AHK_Unknown_Ticket_href_list.txt
      Loop, parse, All_Ticket_Full_HTML, `n, `r
      {
        if (RegExMatch(A_LoopField, ".*href=(.*) data-id", SubPat)) {  ; Each Unknown Ticket Matches
          ;~ MsgBox %A_LoopField%
          StringReplace, SubPat1,SubPat1,", , All    ; Remove Quotes from match (All Occurences)   ; Example  "/MSP/PendingTicketsEdit.aspx??Edit=True&amp;TBL=87&amp;UFAM=&amp;ID=13899258"
          StringReplace, SubPat1,SubPat1,amp;, , All ; Remove amp; from match (All Occurences)     ; Example  /MSP/PendingTicketsEdit.aspx??Edit=True&amp;TBL=87&amp;UFAM=&amp;ID=13899258
                                                                                                   ; Example  /MSP/PendingTicketsEdit.aspx??Edit=True&TBL=87&UFAM=&ID=13899258
          Unknown_Ticket_URL := "https://TicketingSystem.kaseya.com/" . SubPat1
          FileAppend, %Unknown_Ticket_URL%`n ,C:\users\%username%\Documents\AHK_Unknown_Ticket_href_list.txt
          ;~ MsgBox, %Unknown_Ticket_URL%
          ticket_count += 1
        }
      }
      FileRead, href_list, C:\users\%username%\Documents\AHK_Unknown_Ticket_href_list.txt
      Loop, parse, href_list, `n, `r
      {
        if ( InStr(A_LoopField, "/MSP/Pending") ) {
          UT_Session.NewTab()
          UT_Session.Navigate(A_LoopField)
          All_Ticket_Text := UT_Session.querySelector("#ctl00_phContent_txtTicketTitle").value          ; Gets the content of Title Field
          All_Ticket_Text .= UT_Session.querySelector("#ctl00_phContent_txtDetailsWrapper").innerHTML    ; Appends the HTML content of the Details
          Ticket_Title := UT_Session.querySelector("#ctl00_phContent_txtTicketTitle") ; getting  Ticket Title  element
          account := Ticket_AccountLookup(All_Ticket_Text)
          if (account) {  ;
            Ticket_Title_Length := StrLen(Ticket_Title.value)    ; Get Length of Ticket Title
            if (Ticket_Title_Length > 195 ) {   ; Trim Title  -- TicketingSystem Max is 200 Chars
              Ticket_Title.value := SubStr(Ticket_Title.value, 1, 195)   ; Updates Title
            }
            UT_Session.querySelector("#ctl00_phContent_ddlClients_Input").click()   ; Clicks in account box
            UT_Session.querySelector("#ctl00_phContent_ddlClients_Input").SendKey(account . key.class )  ;  Send Account that was matched in      account := Ticket_AccountLookup(All_Ticket_Text)
            sleep 1000           ;   Allow time for dropdown box to populate after typing account
            UT_Session.querySelector(".rcbItem").click()   ;   Click the account listed
            sleep 1000    ; Allow form to autofill after clicking the account
            UT_Session.querySelector("#ctl00_phMenu_btnSave").click()  ; Click  Save
          } else { ; Account is not defined unable to update ticket
            FileAppend, %All_Ticket_Text%`n-_-_-_`n ,C:\users\%username%\Documents\AHK_Log_Unknown_Tickets_Not_Resolved.txt   ; Log unmatched tickets text contents
          }
        }
      }
      Loop, Files, C:\Users\%username%\AppData\Local\Temp\*.*, D
      {
        if InStr(A_LoopFileFullPath, "scoped_dir") || InStr(A_LoopFileFullPath, "edge_BITS_") {   ; Remove Edge Specific Temp Files
          FileRemoveDir, % A_LoopFileFullPath, 1
        }
      }
      flash_Screen_Rainbow()
      ;~ Run cmd.exe /c taskkill /f /t /im:msedgedriver.exe , , Min    ; Kills the web driver  /t flag  also kills open browsers launched from webdriver
      return
;}
    l:: ; {F4} -- Unknown Ticket Automation -- Infinite Loop
;{   Unknown Ticket Automation   Loop Mode
    MSEdge_UT := new Rufaydium("msedgedriver.exe","--port=54326")
      ;~ MsgBox % User_Config_File.Rufaydium.HeadlessMode
    if (User_Config_File.Rufaydium.HeadlessMode) {
      MSEdge_UT.capabilities.HeadlessMode := true
    }
    UT_Session := MSEdge_UT.NewSession()    ; Launch Web Browser with specified capabilities
    MSEdge_UT.Navigate("portal.office.com")
      Sleep User_Config_File.Rufaydium.O365_IsUserAuthenticated_Sleep
      If InStr(MSEdge_UT.url, "login.microsoftonline.com") {
        Data = ;nul
        Data := CredRead("Office365_cred")
        If (Data.name = "")
        {
          flash_Screen("Red")
        } else {
          f3q_ClipB := Data.username
          f3a_ClipB := Data.password
          clipboard = %f3q_ClipB%
          flash_Screen("Blue")
          DisplayTooltip("Personal Credential Read`n" . Office365_cred,User_Config_File.Timer.UsernameCopied_Tooltip,1)
        }
        MsgBox, Please Login To Office Before Clickng Ok
      }
    Loop, Files, C:\Users\%username%\AppData\Local\Temp\*.*, D
      {
        if InStr(A_LoopFileFullPath, "scoped_dir") || InStr(A_LoopFileFullPath, "edge_BITS_") {   ; Remove Edge Specific Temp Files
          ;~ DisplayTooltip(A_LoopFileFullPath,1000,1)
          FileRemoveDir, % A_LoopFileFullPath, 1
        }
      }
    Loop   ; Infinite Loop
    {
      ;{   Authenticate to TicketingSystem
      UT_Session.Navigate("https://TicketingSystem.kaseya.com/Gateway.aspx?enc=%2fZyI3iknrKMpB8K8w6gKLbUeJsnECRcITlan3qIkV9GLYTVtQkVPNqWJY%2ftQiYkl%2fATNIRE2wODGGweAsGY%2fDQ%3d%3d")   ; TicketingSystem GateWay
      UT_Session.querySelector("#ctl00_MainContent_txtUserID").value := "cjohnson"   ; Username to login as
      UT_Session.querySelector("#ctl00_MainContent_btnNext").click()    ; Click Button to login
      ;}
      UT_Session.Navigate("https://TicketingSystem.kaseya.com/MSP/PendingTickets.aspx")    ;  Open "Unknown Tickets" (Root URL)
      All_Ticket_tbody := UT_Session.querySelector("#ctl00_phListing_ctrlListing_dgrdListing_ctl00 > tbody")  ; Unkown Ticket Listing (Parent Element)
      All_Ticket_Full_HTML := All_Ticket_tbody.InnerHTML     ; Grab the HTML from the unknown tickets
      ;~ RegExMatch(All_Ticket_Full_HTML, ".*href=(.*) data-id", SubPat)  ;
      ticket_count = 0
      assigned_count = 0
      FileDelete, C:\users\%username%\Documents\AHK_Unknown_Ticket_href_list.txt
      Loop, parse, All_Ticket_Full_HTML, `n, `r
      {
        if (RegExMatch(A_LoopField, ".*href=(.*) data-id", SubPat)) {  ; Each Unknown Ticket Matches
          ;~ MsgBox %A_LoopField%
          StringReplace, SubPat1,SubPat1,", , All    ; Remove Quotes from match (All Occurences)   ; Example  "/MSP/PendingTicketsEdit.aspx??Edit=True&amp;TBL=87&amp;UFAM=&amp;ID=13899258"
          StringReplace, SubPat1,SubPat1,amp;, , All ; Remove amp; from match (All Occurences)     ; Example  /MSP/PendingTicketsEdit.aspx??Edit=True&amp;TBL=87&amp;UFAM=&amp;ID=13899258
                                                                                                   ; Example  /MSP/PendingTicketsEdit.aspx??Edit=True&TBL=87&UFAM=&ID=13899258
          Unknown_Ticket_URL := "https://TicketingSystem.kaseya.com/" . SubPat1
          FileAppend, %Unknown_Ticket_URL%`n ,C:\users\%username%\Documents\AHK_Unknown_Ticket_href_list.txt
          ;~ MsgBox, %Unknown_Ticket_URL%
          ticket_count += 1
        }
      }
      FileRead, href_list, C:\users\%username%\Documents\AHK_Unknown_Ticket_href_list.txt
      Loop, parse, href_list, `n, `r
      {
        if ( InStr(A_LoopField, "/MSP/Pending") ) {
          UT_Session.Navigate(A_LoopField)
          All_Ticket_Text := UT_Session.querySelector("#ctl00_phContent_txtTicketTitle").value          ; Gets the content of Title Field
          All_Ticket_Text .= UT_Session.querySelector("#ctl00_phContent_txtDetailsWrapper").innerHTML    ; Appends the HTML content of the Details
          Ticket_Title := UT_Session.querySelector("#ctl00_phContent_txtTicketTitle") ; getting  Ticket Title  element
          account := Ticket_AccountLookup(All_Ticket_Text)
          if (account) {  ;
            assigned_count += 1
            Ticket_Title_Length := StrLen(Ticket_Title.value)    ; Get Length of Ticket Title
            if (Ticket_Title_Length > 195 ) {   ; Trim Title  -- TicketingSystem Max is 200 Chars
              Ticket_Title.value := SubStr(Ticket_Title.value, 1, 195)   ; Updates Title
            }
            UT_Session.querySelector("#ctl00_phContent_ddlClients_Input").click()   ; Clicks in account box
            UT_Session.querySelector("#ctl00_phContent_ddlClients_Input").SendKey(account . key.class )  ;  Send Account that was matched in      account := Ticket_AccountLookup(All_Ticket_Text)
            sleep 1000           ;   Allow time for dropdown box to populate after typing account
            UT_Session.querySelector(".rcbItem").click()   ;   Click the account listed
            sleep 1000    ; Allow form to autofill after clicking the account
            UT_Session.querySelector("#ctl00_phMenu_btnSave").click()  ; Click  Save
          } else { ; Account is not defined unable to update ticket
          }
        }
      }
      text := "Unknown - " . ticket_count
      if (ticket_count) {
        text .= "`nAssigned - " . assigned_count
      }
      DisplayTooltip(text,User_Config_File.Timer.F4_L_Unknown_Ticket_Loop_ToolTip,1)
      sleep User_Config_File.Timer.F4_L_Unknown_Ticket_Loop_Cooldown
    }
    return    ; This wont ever happen... Gotta kill with Escape
;}
}
#If
;}
;}

F6::  return  ; zzz_ Control Character for shortcuts - Developer Stuff
{
#If (A_PriorHotKey = "F6" AND A_TimeSincePriorHotkey < User_Config_File.Timer.FN_Key_Timer)
;{   ; {F6} -- DEVELOPMENT USE   ^w  !w     --- Gets ( Window Name / ahk HWND Window Handle) of window under mouse
    ^w:: ; {F6} -- Gets name of window under mouse cursor
       MouseGetPos,xpos,ypos,cursWin,cursControl
       WinGetTitle, cursTitle, ahk_id %cursWin%
       DisplayTooltip(cursTitle,3000,1)
       w_ClipB = %cursTitle%
       return
    !w:: ; {F6} -- Gets name of ahk HWND (window handle) under mouse cursor
       MouseGetPos,xpos,ypos,cursWin,cursControl
       WinGetTitle, cursTitle, ahk_id %cursWin%
       DisplayTooltip(cursWin,3000,1)
       s_ClipB = %cursTitle%
       return
;}
}
#If

F9:: return ; azzz_ Control Character for shortcuts - User Config File Management
{
#If (A_PriorHotKey = "F9" AND A_TimeSincePriorHotkey < User_Config_File.Timer.FN_Key_Timer)
    !0:: ; {F9} -- User Configuration - Reset User Configuration to defaults
      Reset_INI_File_From_Template(User_Config_File_Path,User_Config_Template_Path)
      reload_Function()
      return
    e:: ; {F9} -- User Configuration - Edit User Configuration
      Run, %User_Config_File_Path%,,,PID    ; Launch User Config File in default editor
      Process, WaitClose, %PID%   ; Wait for user to close config file
      User_Config_File := Ini(User_Config_File_Path)   ; Read config
      flash_Screen("Yellow")
      return
!+e:: ; {F9} -- DEVELOPER USE --- Edit Script Template --- PLZ DONT TOUCH TNX
      Run, %MDrive_User_Config_Template_Path%    ; Launch User Config File in default editor
      return
#If
}

;}     F1 - F9  (Collapse)
;{ Mouse Hotkeys

;{  Alt Shift  ---  Click Tracking Functions
!+MButton::  ; Mouse - display tooltip of relative mouse cordinates. 0,0 is upper left of active window
   MouseGetPos, xpos, ypos
   mousePosition = %xpos%, %ypos%`n
   DisplayTooltip(mousePosition,3000,1)
   return

!+LButton::  ; Mouse - Click and Track relative mouse cordinates.  Recall with !+RButton
   MouseGetPos, xpos, ypos
   mousePosition = Click, %xpos%, %ypos%`n
   DisplayTooltip(mousePosition,3000,1)
   Click
   total_clicklist = %total_clicklist%%mousePosition%
   return

!+RButton::  ;  Mouse - Recalls tracked clicks from   !+LButton
   clipboard = %total_clicklist%
   DisplayTooltip(total_clicklist,3000,1)
   return
;}

;{ Credentail Manager --- Copy Username / Password
^WheelUp::  ; Mouse Scroll - Copies Username of last looked up credential
    if (Data) {
      clipboard := Data.username
      DisplayTooltip(clipboard,User_Config_File.Timer.UsernameCopied_Tooltip,1)
    } else {
      Data := CredRead(Password Management System_Credential)
      if (Data) {
        clipboard := Data.username
      } else {
        DisplayTooltip("No Password Management System Credential Tagged...`nTag a Credential with Ctrl MiddleMouse",2000,1)
      }
    }
    return

^WheelDown:: ; Mouse Scroll - Copies Password of last looked up credential
    if (Data) {
      clipboard := Data.password
      if ( (InStr(Data.name, "MY_COMPANY_NAME_AD_cred")) || (InStr(Data.name, "Office365_cred")) || (InStr(Data.name, "company123_Admin_cred")) || (InStr(Data.name, "EmployeeTimeTrackingSoftware_cred"))) {
        DisplayTooltip(Data.name . "  --  Password Copied",User_Config_File.Timer.PasswordCopied_Tooltip,1)
      } else {
        DisplayTooltip(clipboard,2000,1)
      }
    } else {
      Data := CredRead(Password Management System_Credential)
      if (Data) {
        clipboard := Data.username
      } else {
        DisplayTooltip("No Password Management System Credential Tagged...`nTag a Credential with Ctrl MiddleMouse",2000,1)
      }
    }
    return
;}

;}

