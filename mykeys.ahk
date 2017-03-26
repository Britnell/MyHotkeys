
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases. 
; #Warn  ; Enable warnings to assist with detecting common errors. 
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability. 
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory. 

;	#	#	Hotkey Reference
; +	shift
; # 	windows
; !	alt
; ^	control
; &	custom combination hotkeys

; * Long Functions first
; * scroll down for Hotkeys
; * Some debug / test functions at the end

;	#	#	#	#	#
; 	#	# 	Functions	#
;	#	#	#	#	#

;	This function finds and switches to an open window which INCLUDES the given TITLE
;	Input : Window title - e.g. Chrome, Firefox
;	
SwitchToWin(TheWindowTitle) 
{ 
SetTitleMatchMode,2 
DetectHiddenWindows, Off 
IfWinExist, %TheWindowTitle% 
{ 
WinGet, winid, ID, %TheWindowTitle% 
DllCall("SwitchToThisWindow", "UInt", winid, "UInt", 1) 
} 
Return 
} 

;	This function finds and switches to an open window which INCLUDES the given TEXT
;	Input : Text inside the window
;	
SwitchToWinText(TheWindowText) 
{ 
SetTitleMatchMode,2 
DetectHiddenWindows, Off 
   ;## WinGet, OutputVar [, Cmd, WinTitle, WinText, ExcludeTitle, ExcludeText]
WinGet, winid, ID,,%TheWindowText% 
DllCall("SwitchToThisWindow", "UInt", winid, "UInt", 1) 
Return 
}  

;	This function creates a hyperlink inside OneNote, 
;		but could be adapted to do the same in Word or other
;	It takes address of current tab and website title found in Window
;
O_N_MakeHyperlink_tab()
{
;go chrome
SwitchToWin("Chrome")
clipboard = 	; empty
Sleep 10
WinGetActiveTitle, Title 	; Get page title from window title
Send, ^l		; Go link bar
Sleep 10
Send, {Esc}		; Show page address (in case of alterations)
Sleep 5
Send, ^c		; Copy link
ClipWait
Send, {TAB}		; Switch focus out of address field
SwitchToWin("OneNote")
Send, {End}		; Go to end of line and start new line 
Send, {Enter}
Sleep 10
Send, %Title%	; Enter Page title
Sleep 10
Send, +{Home}	; Mark text
Sleep 5
Send, ^k		; Make hyperlink
Sleep 10
Send, %clipboard%	;(^v) Enter address
Sleep 10
Send, {Enter}		; OK
Sleep 10
Send, {End}			; Exit line
Send, {Enter}
Return
}

;	Creates hyperlink in chrome, address of current tab, with title you highlight in the page
;		When in Chrome, mark a few words / line
;		then trigger hotkey
;
O_N_MakeHyperlink_copy()
{
;if in chrome		;SwitchToWin("Chrome")
clipboard = 			; empty clipboard
Send, ^c			; Copy selected title
ClipWait			; waits for copy
;clipboard = %clipboard%		; unformats, but leaves new lines
Send, ^l			; Go to address bar
Sleep 5				; pasting and copying title in address bar removes new lines in text
Send, ^v			; %clipboard% , paste	
Sleep 10
Send, ^a
clipboard = 			
Send, ^c			; Copy unformated title
ClipWait
title = %clipboard%
Send, {Esc}			; reveal page address
Sleep 5
clipboard = 
Send, ^c			; Copy link
ClipWait
address = %clipboard%
Send, {Tab}			; leave bar
SwitchToWin("OneNote") 	
Sleep 10
Send, {End}			; Got to end of line and start new line 
Send, {Enter}			; NL
Sleep 5
Send, %title%			; Paste title
Sleep 10
Send, +{Home}			; Mark text
Sleep 10
Send, ^k				; Make hyperlink
Sleep 10
Send, %address%			; enter address
Sleep 20
Send, {Enter}			; ok
Sleep 10
Send, {End}			; Exit line
Send, {Enter}
Return
}

;		Creates hyperlink in OneNote 
;			Copies address of link in Chrome the mouse is overing over 
;			( may only work in Chrome / my version. Depends on right-click menu)
;			stops at link text field for you to manually enter link name
;
O_N_MakeHyperlink_link()
{
clipboard = 
KeyWait, Alt		; My hotkey uses Alt+... thus script needs to WAIT for ALT to be released
Send {RButton}		; Right click
Sleep 10
Send {Down 5}		; > Copy link location
Send {Enter}
ClipWait
SwitchToWin("OneNote") 	
Sleep 10
Send {End}			; Go to new line
Send {Enter}
Sleep 5
Send ^k				; Create hyperlink
Sleep 10
Send +{Tab 3}		; Go to address field
Send %clipboard%		; Paste
Send +{Tab}			; Go to text field
Return				; and exit
}


;	#	#	#	#	#	#
; 	#	# 	Window switcher #	#
;	#	#	#	#	#	#

;	Press Alt+'E' to switch to last open Windows Explorer 
;		Opens new explorer if none active
;		Note : As Explorer window always carries name of Folder 
;		Solution : Windows Text always starts with "Address:C://..."
;			>> Search window by text
!e::
IfWinExist,, Address:
{
SwitchToWinText("Address:") 
}		
Else
{
Send, #e 
}
Return

;	Press Alt + 'R' to switch to Google Chrome
!r::SwitchToWin("Chrome") 

;	Press Alt + 'W' to switch to OneNote
!w::SwitchToWin("OneNote")  

;	Press Alt + 'Q' to switch to Outlook
!q::SwitchToWin("Outlook")  


; 	Tab Switch
;		Alt + X presses Alt+Tab X times, to go to 2nd last / 3rd last window
;		Does not work in all windows
;
!1::Send {Alt down}{TAB 1}{Alt up}
!2::Send {Alt down}{TAB 2}{Alt up}
!3::Send {Alt down}{TAB 3}{Alt up}
!4::Send {Alt down}{TAB 4}{Alt up}
!5::Send {Alt down}{TAB 5}{Alt up}

;	#	#	Capslock	#	#

;	First I changed my capslock key to Right-click / AppsKey
;
;Capslock::AppsKey 

;	Capslock key switches to Google Chrome and Opens New tab
;
Capslock::
SwitchToWin("Chrome")
Sleep 10
IfWinActive, Chrome
{
Send ^t
}
Return 

;	Alt + Capslock is like Chromebook key
;		Mark Text
;		Marked text will be copied and googled in new Chrome tab
;
!Capslock::
clipboard = 	;empty
Send ^c
ClipWait
SwitchToWin("Chrome") 	
Sleep 10
IfWinActive, Chrome		
{
Send ^t
Send ^v
Send {Enter}
}
Return 

;		Windows + 'C' starts calculator
;
#c:: 
Run, calc.exe 
WinWait, Calculator 
Return 

 
;	#	#	#	#	#
;	#	# 	Commands	#
;	#	#	#	#	#
;
;	I use a simple Input - GUI to enter custom word or letter commands

;	To open custom "command-prompt" press the RIGHT Shift KeyWait
;	
RShift::
InputBox, CmdVar, Autohotkey Commandline,, 
if !ErrorLevel		; Only if nothing is entered
{ ; Open If

if ( CmdVar = "w")		; Open Wikipedia
{
	; Open Web page
	SwitchToWin("Chrome") 
	Sleep 10
	IfWinActive, Chrome
	{
	Send ^t
	Sleep 5
	Send https://en.wikipedia.org/wiki/Main_Page
	Send {Enter}
	}
}
else if ( CmdVar = "ck")		; Carry out Copy & Hyperlink function
{
	O_N_MakeHyperlink_copy()
}
else if ( CmdVar = "tk")		; Carry out Current tab Hyperlink function
{
	O_N_MakeHyperlink_tab()
}
else if ( CmdVar = "lk")		; Start new hyperlink with current address
{
	O_N_MakeHyperlink_link()
}
else							; Print Unknown command
{
	; Default
	MsgBox, Go-go: %CmdVar%
}

} ; Close If
Return

;	#	#	Custom hotkeys	#	
;	
;	Hold down 'K' and press 'C' or vice versa
;		~ is important as otherwise keys lose their original function
;		( means no more K or C )

~k & c::O_N_MakeHyperlink_copy()

~k & t::O_N_MakeHyperlink_tab()


;	#	#	#
;	#	Chrome  #
;	#	#	#

;	Keys specific to active window
;		Here I use the ahk_class to identify Google Chrome

;	Key tells you if its pressed in Google Chrome, or any other window
#q::
IfWinActive ahk_class Chrome_WidgetWin_1
{
MsgBox, chrome
}
Else
{
MsgBox, any
}
Return

;	Hold Alt and Left-Mouse-Click on a link in Browserto
;		to trigger OneNote hyperlink function
!LButton::
IfWinActive ahk_class Chrome_WidgetWin_1
{
O_N_MakeHyperlink_link()
}
Return

;	Alt + Up or Down scroll window mouse is above
!Up::WheelUp
!Down::WheelDown

;	In Chrome only : Ctrl + Left or Right changes tab
;		(usually Ctrl + Page-Up or Page-down])
^Left::
IfWinActive ahk_class Chrome_WidgetWin_1
{
Send, ^{PgUp}
}
Else
{
Send, ^{Left}
}
Return

^Right::
IfWinActive ahk_class Chrome_WidgetWin_1
{
Send, ^{PgDn}
}
Else
{
Send, ^{Right}
}
Return


;	#	#	#	#	#
; 	#	# 	Inserts 	#
;	#	#	#	#	#

;	An active listener is triggered if specific keys are pressed in sequence
;		This strangely does not ALWAYS work. 
;		I think Microsoft-Office auto-correct interferes with it

;  Replaces A with B
;  :*SI:A::B

:*SI:@to::my Email
:*SI:@ta::my other email
:*SI:@@::another email
:*SI: u :: you 
:*SI:mfg::Mit freundlichen Grüßen,
:*SI:.ae ::ä
:*SI:.ue ::ü
:*SI:.oe ::ö
:*SI:.eu ::€
:*SI:.p ::£


;	#	#	Debug	#
; 	#	# 	and Test 
;	
#w:: 
WinGetActiveTitle, Title 
MsgBox, The active window is "%Title%". 
Return 


;End 