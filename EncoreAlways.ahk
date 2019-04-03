;	Name:		EncoreAlways
;	Version:	0.9
;	Author:		Lucas Bodnyk
;
;	I based this on the SierraWrapper - they do very similar things.
;	This script draws code from the WinWait framework by berban on www.autohotkey.com, as well as some generic examples.
;	All variables "should" be prefixed with 'z'.
;
;
;	All User Startup is '\\<Machine_Name>\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup'.
;	I recommend placing a shortcut there, pointing to this, but I have no idea where to put this.
;
;
; -----------------------------  Revision History  --------------------------------------------------------------
;
;   03/15/2019  Greg Pruett    Changed script to run Encore in Chrome instead of IE.
;   04/03/2019	Greg Pruett	   Altered script to clean up some variable initialization.  
;
;


#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
#InstallKeybdHook ; necessary for A_TimeIdlePhysical
#InstallMouseHook ; necessary for A_TimeIdlePhysical
OnExit("ExitFunc") ; Register a function to be called on exit
global zLocation := RegExReplace(A_ComputerName, "-[\s\S]*$") ; returns everything before (by replacing everything after) the first dash in the machine name
StringLower, zLocation, zLocation
zNumber := RegExReplace(A_ComputerName, "[\s\S]*-") ; returns everything after (by replacing everything before) the last dash in the machine name
global zNameWithSpaces := A_ComputerName . "                "
StringLeft, zNameWithSpaces, zNameWithSpaces, 15 ; adds whitespace so the log will be justified.
SplitPath, A_ScriptName, , , , ScriptBasename
StringReplace, AppTitle, ScriptBasename, _, %A_SPACE%, All
global zActivityPeriods
global zInSession
global zDisplayingIdleWarning
global zCountDown
global zSessionStart
global zSessionEnd := A_TickCount
global zEncoreLogPath
global zSessionTimeout

;
;	BEGIN INITIALIZATION SECTION
;

Try {
	Log("")
	Log("   EncoreAlways initializing for machine: " A_ComputerName)
} Catch	{
	MsgBox Testing EncoreAlways.log failed! You probably need to check file permissions. I won't run without my log! Dying now.
	ExitApp
}
Try {
	IniWrite, 1, EncoreAlways.ini, Test, zTest
	IniRead, zTest, EncoreAlways.ini, Test, 0
	IniDelete, EncoreAlways.ini, Test, zTest
} Catch {
	Log("!! Testing EncoreAlways.ini failed! You probably need to check file permissions! I won't run without my ini! Dying now.")
	MsgBox Testing EncoreAlways.ini failed! You probably need to check file permissions! I won't run without my ini! Dying now.
	ExitApp
}
IniRead, zClosedClean, EncoreAlways.ini, Log, zClosedClean, 0
IniRead, zEncoreLogPath, EncoreAlways.ini, General, zEncoreLogPath, %A_Space%
IniRead, zName, EncoreAlways.ini, General, zName, %A_Space%
IniRead, zExitPassword, EncoreAlways.ini, General, zExitPassword, %A_Space%
IniRead, zSessionTimeout, EncoreAlways.ini, General, zSessionTimeout, %A_Space%
Log("## zClosedClean="zClosedClean)
Log("## zEncoreLogPath="zEncoreLogPath)
Log("## zName="zName)
Log("## zExitPassword="zExitPassword)
Log("## zSessionTimeout="zSessionTimeout)
zInSession := 0
zDisplayingIdleWarning := 0
If (zClosedClean = 0) {
	Log("!! It is likely that EncoreAlways was terminated without warning.")
	}
If (zEncoreLogPath = "") {
	zEncoreLogPath := A_WorkingDir
	Log("ii I will be logging browser activity locally`, to "zEncoreLogPath)
	}
If (zName = "") {
	zName := zLocation . zNumber
	Log("ii I will be parsing the machine name`, "zName)
	}
If (zExitPassword = "") {
	zExitPassword := ""
	Log("ii The password is blank!")
	}
;zSessionTimeout := (zSessionTimeout * 1000)
;MsgBox %zSessionTimeout%
If (zSessionTimeout = "") {
	zSessionTimeout := 120
	Log("ii No session timeout was specified, using default of 120 seconds.")
	}

	
IniWrite, 0, EncoreAlways.ini, Log, zClosedClean
Log("ii Initialization finished`, starting up...")

ClearSettingsAndRestart()
SetTimer, __Main__, 500 ; Every half second, do a bunch of things.

__Main__:
	If !ProcessExist("chrome.exe")
	{
		Log("ie Starting browser in kiosk mode.")
		Run "chrome.exe" -incognito -kiosk http://dcls-mt.iii.com, "c:\Program Files(x86)\Google\Chrome\Application\"
	}
	;Log("A_TimeIdlePhysical: "A_TimeIdlePhysical)
	If (A_TimeIdlePhysical < 500) {
		zActivityPeriods++
		If (zDisplayingIdleWarning = 1) {
		DestroyIdleWarning()
		FileAppend, %A_YYYY%/%A_MM%/%A_DD% %A_Hour%:%A_Min%:%A_Sec%    %zNameWithSpaces%    <> A timeout was interrupted.`n, %zEncoreLogPath%%zLocation%-encorelog.txt
		Log("<> A timeout was interrupted.")
		}
	}
	If (zActivityPeriods > 6 and zInSession = 0) {	; if we've built up enough activity and we didn't already have an active session
	BeginSession()
	}
	If (A_TimeIdlePhysical > 15000 and zActivityPeriods < 7 and zActivityPeriods > 0 and zInSession = 0) {	; if idle for more than 15 seconds and a session hasn't already been started, but there was a little activity! (phew)
		zActivityPeriods := 0 								; clear the activity tracking variable
		Log(".. Someone touched the computer but did not start a session")
	}	
	If (A_TimeIdlePhysical > 150000 and zInSession = 1 and zDisplayingIdleWarning = 0) { ; If idle for more than 150 seconds and a session was started AND we're not already warning of inactivity
		zDisplayingIdleWarning := 1
		DisplayIdleWarning()
	}
	;Log("zActivityPeriods: "zActivityPeriods)
	return

BeginSession()
{
	zSessionStart := A_TickCount
	zInSession := 1
	zElapsedTime := ROUND((A_TickCount - zSessionEnd)/1000)/60
	;MsgBox %zElapsedTime%
	FileAppend, %A_YYYY%/%A_MM%/%A_DD% %A_Hour%:%A_Min%:%A_Sec%    %zNameWithSpaces%    >> A session has begun. Minutes since last session: %zElapsedTime%`n, %zEncoreLogPath%%zLocation%-encorelog.txt
	Log(">> A session has begun. Minutes since last session: " zElapsedTime)
	zBrowserCleared := 0
}
	
DisplayIdleWarning()
{
	global Secs := 30
	Gui, platypus:New
	Gui +AlwaysOnTop 
	Gui, Font, S24 CWhite, Verdana
	Gui, Color, black, black
	Gui -Caption
	Gui, Add, Text, vzCountDown x0 y0 w%A_ScreenWidth% h%A_ScreenHeight% +Center, `n`n`n`n`n`nThe session will time out in %Secs% seconds.`n Move the mouse to prevent this.
	Gui, Show, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight%
	Gui, platypus: +LastFound  ; Make the GUI window the last found window for use by the line below.
	WinSet, Transparent, 192
	WinSet, AlwaysOnTop, On, GUI ; not sure on syntax for WinSet
	SetTimer, CountDown, 1000
}

DestroyIdleWarning()
{
	Gui platypus: Destroy
	zDisplayingIdleWarning := 0
	SetTimer, CountDown, Off
}

EndSession()
{
	zSessionEnd := A_TickCount
	zInSession := 0
	zElapsedTime := ROUND((A_TickCount - zSessionStart)/1000)
	FileAppend, %A_YYYY%/%A_MM%/%A_DD% %A_Hour%:%A_Min%:%A_Sec%    %zNameWithSpaces%    << A session has ended. Seconds since session start: %zElapsedTime%`n, %zEncoreLogPath%%zLocation%-encorelog.txt
	Log("<< A session has ended. Seconds since session start: " zElapsedTime)
	zActivityPeriods := 0
	ClearSettingsAndRestart()
}

CountDown:
Secs--
If (Secs <1) {
	DestroyIdleWarning()
	EndSession()
}
GuiControl, platypus:,zCountDown, `n`n`n`n`n`nThe session will time out in %Secs% seconds.`n Move the mouse to prevent this.
Return 
	
ClearSettingsAndRestart() {
	BlockInput On
	Gui, New
	Gui +AlwaysOnTop 
	Gui, Font, S36 CDefault, Verdana
	Gui, Color, gray, gray
	Gui -Caption
	Gui, Add, Text, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight% +Center, `n`n`n`n`n`nPlease wait, we're reloading the browser.
	Gui, Show, x0 y0 w%A_ScreenWidth% h%A_ScreenHeight%
	WinSet, AlwaysOnTop, On, GUI ; not sure on syntax for WinSet
	Process, Close, chrome.exe
	
;	Send ^+{DEL} ; send Ctl-Shft-Del to open "Clear Browsing Data" settings window
;    WinWaitActive, Settings - Clear browsing data - Google Chrome, , ; Wait till window is active

;    Send {Tab}{Enter} ; tab to "Clear Browsing Data" button and select
;    WinWaitActive, Settings - Google Chrome, , ; Wait till Clear Browsing Data is complete
	
	
	zBrowserCleared := 1 ; save value to variable so script won't clear browser until it's been used again.
	Run "chrome.exe" -incognito -kiosk http://dcls-mt.iii.com, "C:\Program Files (x86)\Google\Chrome\Application\"
	Sleep 7000
	Gui, Destroy
	Blockinput Off
	return
}
	
F12::
	return ; dirty way to disable developer mode
	
!F4::
	InputBox, zPassword, Password, Enter the password., HIDE, , , , , , 10
	If (zPassword == zExitPassword)
	{
		ExitApp
	}
	return
	
ProcessExist(Name){
	Process,Exist,%Name%
	return Errorlevel
}

ProgressOff:
Progress, Off
Return

; functions to log and notify what's happening, courtesy of atnbueno
Log(Message, Type="1") ; Type=1 shows an info icon, Type=2 a warning one, and Type=3 an error one ; I'm not implementing this right now, since I already have custom markers everywhere.
{
	global ScriptBasename, AppTitle
	IfEqual, Type, 2
		Message = WW: %Message%
	IfEqual, Type, 3
		Message = EE: %Message%
	IfEqual, Message, 
		FileAppend, `n, %ScriptBasename%.log
	Else
		FileAppend, %A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%.%A_MSec%%A_Tab%%Message%`n, %ScriptBasename%.log
	Sleep 50 ; Hopefully gives the filesystem time to write the file before logging again
	Type += 16
	;TrayTip, %AppTitle%, %Message%, , %Type% ; Useful for testing, but in production this will confuse my users.
	;SetTimer, HideTrayTip, 1000
	Return
	HideTrayTip:
	SetTimer, HideTrayTip, Off
	TrayTip
	Return
}
LogAndExit(message, Type=1)
{
	global ScriptBasename
	Log(message, Type)
	FileAppend, `n, %ScriptBasename%.log
	Sleep 1000
	ExitApp
}

ExitFunc(ExitReason, ExitCode)
{
    if ExitReason in Exit
	{

		Process, Close, chrome.exe
		IniWrite, 1, EncoreAlways.ini, Log, zClosedClean
		Log("xx User hit Alt-F4 and correctly entered password`, dying now.")
	}
	if ExitReason in Menu
    {
        MsgBox, 4, , This will kill Encore.`nAre you sure you want to exit?
        IfMsgBox, No
            return 1  ; OnExit functions must return non-zero to prevent exit.
		Process, Close, chrome.exe
		IniWrite, 1, EncoreAlways.ini, Log, zClosedClean
    }
	if ExitReason in Logoff,Shutdown
	{

		Process, Close, chrome.exe
		IniWrite, 1, EncoreAlways.ini, Log, zClosedClean
		Log("xx System logoff or shutdown in process`, dying now.")
	}
		if ExitReason in Close
	{

		Process, Close, chrome.exe
		IniWrite, 1, EncoreAlways.ini, Log, zClosedClean
		Log("!! The system issued a WM_CLOSE or WM_QUIT`, or some other unusual termination is taking place`, dying now.")
	}
		if ExitReason not in Close,Exit,Logoff,Menu,Shutdown
	{
		Process, Close, chrome.exe
		IniWrite, 1, EncoreAlways.ini, Log, zClosedClean
		Log("!! I am closing unusually`, with ExitReason: " ExitReason "`, dying now.")
	}
    ; Do not call ExitApp -- that would prevent other OnExit functions from being called.
}