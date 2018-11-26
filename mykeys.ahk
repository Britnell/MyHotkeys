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


			; WinExist()
	; https://www.autohotkey.com/docs/commands/WinExist.htm

			; WinTitle
	; https://www.autohotkey.com/docs/misc/WinTitle.htm

			; WinGet
	; ## WinGet, OutputVar [, Cmd, WinTitle, WinText, ExcludeTitle, ExcludeText]
	; WinGet, winid, ID,,%TheWindowText% 




;	#	#	#	#	#
; 	#	# 	Functions	#
;	#	#	#	#	#



		;	This function finds and switches to an open window which INCLUDES the given TITLE
		;	Input : Window title - e.g. Chrome, Firefox
	
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


O_N_MakeHyperlink_tab()
{
;go chrome
IfWinActive, Mozilla		
	{
		Sleep 5
	}
	Else
	{
		SwitchToWin("Mozilla")
		WinWait, Mozilla   ;;Sleep 100
	}
clipboard = 
WinGetActiveTitle, Title 	; Get page title from window title
;; Title includes" - Mozilla Firefox" = 18 chars
;;StringTrimRight, Opvar, IpVar, count
StringTrimRight, Title, Title, 18
Send ^l			; Go link bar
Sleep 10
Send {Esc}		; Show page address (in case of alterations)
Sleep 10
Send ^c			; Copy link
ClipWait 
Send {TAB}		; Switch focus out of address field
SwitchToWin("OneNote")
WinWait, OneNote    ;;Sleep 100
Send {End}		; Go to end of line and start new line 
Send {Enter}
Send ^k		; Make hyperlink
Send +{TAB 3}
Send ^v
Send +{TAB}
Send %Title%
Send {Enter}
;SwitchToWin("Mozilla")
Return
}


		;	Creates hyperlink in chrome, address of current tab, with title you highlight in the page
		;		When in Chrome, mark a few words / line
		;		then trigger hotkey

O_N_MakeHyperlink_copy()
{
clipboard =       ;
Sleep 10
Send, ^c			; Copy selected title
Sleep 50 		; NOT ClipWait, then we can use whatever was copied before, and not marked text does not crash it
title = %clipboard%
	;;MsgBox, name is  %title%
clipboard =        ;
	;;MsgBox, clip = %clipboard%
IfWinActive, Mozilla		
	{
		Sleep 5
	}
	Else
	{
		SwitchToWin("Mozilla")
		WinWait, Mozilla
	}
Sleep 10
Send, ^l			; Go to address bar
Sleep 10
Send, {Esc}			; reveal page address
Sleep 10
Send, ^c			; Copy link
ClipWait
address = %clipboard%
Send, {Tab}			; leave bar
SwitchToWin("OneNote") 	
Sleep 100  ;;WinWait, OneNote  ; NOT wait so it does nto crash..
IfWinActive, OneNote
{
	Send, {End}			; Got to end of line and start new line 
	Send, {Enter}			; NL
	Send, ^k
	Send +{TAB 3}		;; address line
	Send ^v
	Send +{TAB}			;; go to title
	Send %title%
	Send {Enter}
}
Else
{
	MsgBox, Hyperlink copy did not work - Wheres oneNote at?
}
Return
}

;		Creates hyperlink in OneNote 
;			Copies address of link in Chrome the mouse is overing over 
;			( may only work in Chrome / my version. Depends on right-click menu)
;			stops at link text field for you to manually enter link name
;
O_N_MakeHyperlink_link()
{
IfWinActive, Mozilla		
	{
		Sleep 5
	}
	Else
	{
		SwitchToWin("Mozilla")
		WinWait, Mozilla
	}
Sleep 10
Send, ^l			; Go to address bar
Sleep 10
Send, {Esc}			; reveal page address
Sleep 10
Send, ^c			; Copy link
ClipWait
address = %clipboard%
Send, {Tab}			; leave bar
SwitchToWin("OneNote") 	
Sleep 100  ;;WinWait, OneNote  ; NOT wait so it does nto crash..
IfWinActive, OneNote
{
	Send, {End}			; Got to end of line and start new line 
	Send, {Enter}			; NL
	Send, ^k
	Send +{TAB 3}		;; address line
	Send ^v
	Send +{TAB}			;; go to title
	;; and leave me there
	;;Send %title%
	;;Send {Enter}
}
Else
{
	MsgBox, Hyperlink copy did not work - Wheres oneNote at?
}
Return
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
!d::SwitchToWin("Mozilla") 

!a::SwitchToWin("Adobe Acrobat")

	;	Press Alt + 'W' to switch to OneNote
!w::SwitchToWin("OneNote")  

	;	Press Alt + 'Q' to switch to Outlook
!q::SwitchToWin("Outlook")  

;;;!d::SwitchToWin("Notepad")

!s::SwitchToWin("Sublime")

!p::Send Serial.print



		;	Alt + F to search google for Filetype pdf
		;!f::
		;SwitchToWin("Mozilla")
		;Sleep 50
		;Send ^t
		;Sleep 150
		;Send g filetype:pdf{Space}
		;Sleep 20
		;Send ^v{Enter}
		;Return

		;	Alt + Y to search Youtube
		;!y::
		;SwitchToWin("Mozilla")
		;Sleep 50
		;Send ^t
		;Sleep 150
		;Send y{Space}
		;Sleep 20
		;Send ^v{Enter}
		;Return


; 	Tab Switch
;		Alt + X presses Alt+Tab X times, to go to 2nd last / 3rd last window
;		Does not work in all windows
;
;!1::Send {Alt down}{TAB 1}{Alt up}
;!2::Send {Alt down}{TAB 2}{Alt up}
;!3::Send {Alt down}{TAB 3}{Alt up}
;!4::Send {Alt down}{TAB 4}{Alt up}
;!5::Send {Alt down}{TAB 5}{Alt up}

;	#	#	Capslock	#	#

;	First I changed my capslock key to Right-click / AppsKey
;
;Capslock::AppsKey 

;	Capslock key switches to Google Chrome and Opens New tab
;
Capslock::
IfWinExist, ahk_class MozillaWindowClass
{
	SwitchToWin("Mozilla")
	WinWait, Mozilla
	Send ^t
}
Return 

;	Alt + Capslock is like Chromebook key
;		Mark Text
;		Marked text will be copied and googled in new Chrome tab

!Capslock::
Send ^c
;ClipWait
Sleep 20
IfWinExist, ahk_class MozillaWindowClass
{
	SwitchToWin("Mozilla") 	
	WinWait, Mozilla
	Send ^t
	Sleep 10
	Send ^v
	Sleep 10
	Send {Enter}	
}
Return


^Space::Send {SPACE 4}

;		Windows + 'C' starts calculator

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

	if( CmdVar = "p" )		; PDF
	{
		SwitchToWin("Mozilla") 
		Sleep 50
		Send ^t
		Sleep 100
		Send g filetype:pdf{Space}
		Sleep 5
		Send ^v{Enter}
			;SendRaw g filetype:pdf{A_Space}
	}
	else if ( CmdVar = "c")		
	{
		;; make hyperlink with marked text
		O_N_MakeHyperlink_copy()
	}
	else if ( CmdVar = "t")		
	{
		;; make hyperlink with current tab name
		O_N_MakeHyperlink_tab()
	}
	else if ( CmdVar = "k")		
	{
		;; make hyperlink with tab address and leave for user to enter name
		O_N_MakeHyperlink_link()
	}
	else							; Print Unknown command
	{
		; Default
		;;MsgBox, Go-go: %CmdVar%
		SwitchToWin("Mozilla") 
		Sleep 50
		Send ^t
		Sleep 50
		Send %CmdVar%
		Send {Enter}
	}

} ; Close If
Return


				;	#	#	Custom hotkeys	#	
				;	Hold down 'K' and press 'C' or vice versa
				;		~ is important as otherwise keys lose their original function
				;		( means no more K or C )
				;~k & c::O_N_MakeHyperlink_copy()
				;~k & t::O_N_MakeHyperlink_tab()


;   ######################



			;	#	#	#
			;	#	Chrome  #
			;	#	#	#



;	Keys specific to active window
;		Here I use the ahk_class to identify Google Chrome


;	Alt + Up or Down scroll window mouse is above
!Up::WheelUp
!Down::WheelDown

;	In Chrome only : Ctrl + Left or Right changes tab
;		(usually Ctrl + Page-Up or Page-down])
;^Left::Send, ^{PgUp}
;^Right::Send, ^{PgDn}



		;	**   use .,  ,. to scroll between tabs, browser style

:*SI:,.::^{PgDn}
:*SI:.,::^{PgUp}



			;	#	#	#	#	#
			; 	#	# 	Inserts 	#
			;	#	#	#	#	#



;	An active listener is triggered if specific keys are pressed in sequence
;		This strangely does not ALWAYS work. 
;		I think Microsoft-Office auto-correct interferes with it

;  Replaces A with B
;  :*SI:A::B

:*SI:@to::email1
:*SI:@ta::email2
:*SI:@@::email3
:*SI: u :: you 
:*SI:.ae::ä
:*SI:.ue::ü
:*SI:.oe::ö
:*SI:.eu ::€
:*SI:.p ::£


;	#	#	Debug	#
; 	#	# 	and Test 
;	
;#w:: 
;WinGetActiveTitle, Title 
;MsgBox, The active window is "%Title%". 
;Return 


		; ##################
		; # 	
		; # 		Resizes current active window, moves closest corner to mouse cursor

^!w::
WinGetActiveTitle, Title 
WinGetPos X, Y, W, H, %Title%	
MouseGetPos, xpos, ypos 
;MsgBox, The active window is at %X% / %Y%  : %W% x %H% , mouse at %xpos% / %ypos%
if( xpos > W/2){
	; Right side
	if(ypos > H/2){
		WinMove %Title%,, ,, %xpos%, %ypos%
	}
	else {
		;;MsgBox, Top-right! window is at %X% / %Y%  W:H: %W% x %H% , mouse at %xpos% / %ypos%
		; stretch X
		W := xpos
		; move and stretch Y
		Y := Y +ypos
		H := H -ypos
		WinMove %Title%,, %X%,%Y%, %W%, %H%
	}
}
else {
	;; Left side
	if(ypos > H/2){
		;;MsgBox, bottom-left corner
		X := X +xpos
		W := W -xpos
		H := ypos
		WinMove %Title%,, %X%,%Y%, %W%, %H%
	}
	else {
		;;MsgBox, Top-left corner
		X := X +xpos
		Y := Y +ypos
		W := W -xpos
		H := H -ypos
		WinMove %Title%,, %X%,%Y%, %W%, %H%
	}
}
Return



		; ##################
		; # 	
		; # 		Moves current active window, moves window bar under  mouse


^!d::
WinGetActiveTitle, Title 
WinGetPos X, Y, W, H, %Title%	
MouseGetPos, xmos, ymos 
;;MsgBox, The active window is at %X% / %Y%  the mouse at %xmos% / %ymos%
xmos := xmos + X -(W/2)		;;- (W/2)
ymos := ymos + Y -20		;;- 10
;;MsgBox, Next at %xmos% / %ymos%
;;WinMove %Title%,, (%X% +%xmos%), (%Y% +%ymos%)
WinMove %Title%,, %xmos%, %ymos%
Return



;End 