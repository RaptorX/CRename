/*
 * =============================================================================================== *
 * Author           : RaptorX   <graptorx@gmail.com>
 * Script Name      : CRename
 * Script Version   : 1.7
 * Homepage         : -
 *
 * Creation Date    : September 18, 2010
 * Modification Date: February 10, 2013
 *
 * Description      :
 * ------------------
 *
 * -----------------------------------------------------------------------------------------------
 * License          :           Copyright ©2010 RaptorX <GPLv3>
 *
 *          This program is free software: you can redistribute it and/or modify
 *          it under the terms of the GNU General Public License as published by
 *          the Free Software Foundation, either version 3 of the License, or
 *          (at your option) any later version.
 *
 *          This program is distributed in the hope that it will be useful,
 *          but WITHOUT ANY WARRANTY; without even the implied warranty of
 *          MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *          GNU General Public License for more details.
 *
 *          You should have received a copy of the GNU General Public License
 *          along with this program.  If not, see <http://www.gnu.org/licenses/gpl-3.0.txt>
 * -----------------------------------------------------------------------------------------------
 *
 * [GUI Number Index]
 *
 * GUI 01 - Main [CRename]
 *
 * =============================================================================================== *
 */

;+--> ; ---------[Directives]---------
#NoEnv
#SingleInstance Force
; --
SetBatchLines -1
SendMode Input
SetWorkingDir %A_ScriptDir%
;-

;+--> ; ---------[Basic Info]---------
s_name      := "CRename"                ; Script Name
s_version   := "1.7"                    ; Script Version
s_author    := "RaptorX"                ; Script Author
s_email     := "graptorx@gmail.com"     ; Author's contact email
;-

;+--> ; ---------[General Variables]---------
sec         :=  1000                     ; 1 second
min         :=  sec * 60                ; 1 minute
hour        :=  min * 60                ; 1 hour
; --
SysGet, mon, Monitor                    ; Get the boundaries of the current screen
SysGet, wa_, MonitorWorkArea            ; Get the working area of the current screen
mid_scrw    :=  a_screenwidth / 2       ; Middle of the screen (width)
mid_scrh    :=  a_screenheight / 2      ; Middle of the screen (heigth)
; --
s_ini       :=                          ; Optional ini file
s_xml       :=                          ; Optional xml file
;-

;+--> ; ---------[User Configuration]---------
iview       := "i_view32.exe"
depdate     := "dept|day|month|year"
olddepdate  := depdate . "|name"
anclist     := olddepdate . "|preview|SAC|SAP|GB_NN|FileList|press|LockDept"
current     := 1
previous    := 1
;-

;+--> ; ---------[Main]---------
onMessage(0x202, "WM_LBUTTONUP")
if !FileExist("empty.png")
{
	FileInstall, res\empty.png, empty.png
	FileInstall, res\i_view32.exe, i_view32.exe
}
; RegWrite, REG_SZ, HKCR, SystemFileAssociations\image\shell\CRename\command,,"%a_scriptfullpath%" "`%1"

; if 0 > 0            ; A parameter was passed
	; FileAppend, % %1%, % a_temp . "\files"
hpic := wa_Bottom-50-10
lvsz := hpic -195
Gui, +Resize +MaximizeBox +MinSize
Gui, add, Picture, w-1 h%hpic% y10 +Border vpic, % "empty.png"
; Gui, add, GroupBox, w230 h190 x+10 vGB_NN, % "New Name"
Gui, add, Edit, w40 x+10 yp cred Section Limit2 Number gPreview vdept, % 00
Gui, add, Text, x+10 yp+3, % "Department"
Gui, add, Text, cblue x+10, % "{ Tab } to switch"
Gui, add, Edit, cred w40 xs y+10 Limit2 Number gPreview vday, % "day"
Gui, add, Edit, cred w40 x+5 Limit2 Number gPreview vmonth, % "month"
Gui, add, Edit, cred w40 x+5 Limit2 Number gPreview vyear, % "year"
Gui, add, Text, x+10 yp+3, % "Date"
Gui, add, Edit, w200 xs y+10 Limit90 -WantReturn gPreview vname, % "Name"
Gui, add, Text, w200 r2 y+10 center wrap vPreview, % "Preview: Invalid"
Gui, add, Text, cblue center w200 y+10, % "{ + } to Magnify"
Gui, add, Button, w100 h25 y+10 disabled gPrevious vSAP, % "Same && Next &Page"
Gui, add, Button, w100 h25 x+10 +Default gContinue vSAC, % "Save && &Continue"
Gui, add, Text,cblue center w100 xs disabled vpress,% "{ * }"
Gui, add, Text,cblue center w100 x+10,% "{ Enter }"
; Gui, add, Text, cred xs vpress,% "{ - } to break cycle"
Gui, add, ListView, w230 h%lvsz% xs y+15 Grid +Sort AltSubmit -hdr gLV_Sub vFileList, % "File Name"
Gui, Show,, % "CRename"

GuiControl,Focus,dept
Send, {Home}+{End}
return
;-

;+--> ; ---------[Labels]---------
GuiSize:
 anchor("pic", "hw.5x.25", True)
 anchor("FileList", "h", True)
 Loop, Parse, anclist, |
	anchor(a_loopfield, "x", True)
return

GuiDropFiles:
 Gui, ListView, FileList
 addfiles("drop")
return

LV_Sub:
LV_Modify(current, "Select")
if (a_eventinfo && a_guievent = "normal")
{
LV_GetText(tempprev, a_eventinfo)
GuiControl,, pic, % file_dir . "\" . tempprev

GuiControl,Focus,dept

Send, {Home}+{End}
}
return

Previous:
 if (!new_name)
 {
	MsgBox,16,% "Error", % "You cannot start a ""same as previous"" cycle without creating a new name first."
	return
 }

 if (inStr(new_name, "jpeg"))
	oldext := substr(new_name, -3)
 else
	oldext := substr(new_name, -2)

 previous += 1
 char :=
 if InStr(new_name,".jpeg")
	new_name := RegexReplace(new_name := RegexReplace(new_name,"\d+of\d+[^\s]",""), "\.\w{4}", "")
 else
	new_name := RegexReplace(new_name := RegexReplace(new_name,"\d+of\d+[^\s]",""), "\.\w{3}", "")
 new_name := RegexReplace(new_name, "\s?" . ext . "\s?","")
 new_name := RegexReplace(new_name, "\s?" . oldext . "\s?","")
 Loop, Parse, olddepdate, |
	GuiControl,, % a_loopfield, % old%a_loopfield%
 start()
 gdisable("disable")
 GuiControl,, Preview, % preview := "Preview: " new_name . a_space . previous . "of" . previous . "." ext
 Loop, % previous-1
 {
	Loop, % previous-1
	{
		old := file_dir . "\" . new_name . a_space . a_index . "of" . previous-1 . "." . oldext
		new := file_dir . "\" . new_name . a_space . a_index . "of" . previous . "." . oldext
		FileMove, %old%, %new%
		old := file_dir . "\" . new_name . a_space . a_index . "of" . previous-1 . "." . ext
		new := file_dir . "\" . new_name . a_space . a_index . "of" . previous . "." . ext
		FileMove, %old%, %new%
	}
	old := file_dir . "\" . new_name . "." . oldext
	new := file_dir . "\" . new_name . a_space . previous-a_index . "of" . previous . "." oldext
	FileMove, %old% , %new%

	LV_GetText(getext,current-a_index)
	if (inStr(getext, "jpeg"))
		ext := substr(getext, -3)
	else
		ext := substr(getext, -2)
	LV_Modify(current-a_index,"", new_name . a_space . previous-a_index . "of" . previous . "." ext)
 }
 LV_Modify(0, "-Select")
 Sleep, 10
 LV_Modify(current, "Select")
 Gosub, Continue
return

Continue:
	GuiControl, enable, SAP
	GuiControl, enable, press
	
	old_name := new_name
	new_name := RegexReplace(preview, "Preview:\s", "")
	if (!file_dir || !new_name || dept = "XX" || day = "XX" || month = "XX" || year = "XX")
	{
		new_name := old_name
		return
	}
	current += 1
	char :=
	Loop, Parse, olddepdate, |
		old%a_loopfield% := %a_loopfield%
	old_name:= file_dir . "\" . cur_file
	if FileExist(file_dir . "\" . new_name)
	{
		current -= 1
		Msgbox,64,% "File exists", % "The file name you specified already exist. Please type another name"
		GuiControl, Focus, name
		Send, {Home}+{End}
		return
	}
	FileMove, %old_name%, %file_dir%\%new_name%
	addfiles("continue")
return

Preview:
	Gui, Submit, NoHide
	preview := ""
	if (RegexMatch(%A_GuiControl%, "\d{2}") || (a_guicontrol = "dept" && %a_guicontrol%  == ""))
	{
		Gui, Font
		GuiControl, Font, % A_GuiControl
		send {tab}
	}
	else if (A_GuiControl != "name")
	{
		preview := "Invalid"
		Gui, Font, cRed
		GuiControl, Font, % A_GuiControl
	}

	if (a_guicontrol = "dept" && %a_guicontrol%  != "" && !RegExMatch(%a_guicontrol%, "(0[1-9]|1[0-2]|99)") )
	{
		preview := "Invalid"
		Gui, Font, cRed
		GuiControl, Font, % A_GuiControl
	}


	Loop, Parse, depdate, |
	{
		var := %a_loopfield%
		if a_loopfield = %var%
			%a_loopfield% := "XX"
	}

	line :=
	Loop, Parse, name, %a_space%
	{
		if a_loopfield is upper
			line .= a_space a_loopfield
		else 
			line .= format("{:T}",a_loopfield) A_Space
	}

	preview := % "Preview: " (InStr(preview,"invalid") ? preview 
													   : (dept ? format("{:02}",dept) " - " : "") format("{:02}",year) "-" format("{:02}",month) "-" format("{:02}",day) " - " trim(line) "." ext)

	GuiControl,, Preview, % preview
return

2GuiEscape:
2GuiClose:
GuiClose:
 if a_gui = 2
 {
	Gui, 2: Destroy
	toggle :=
 }
 else
	Exitapp
return
;-

;+--> ; ---------[Functions]---------
Anchor(c, a, r = false) { ; v3.5.1 - Titan
	/*
	Function: Anchor
		Defines how controls should be automatically positioned relative to the new
		dimensions of a window when resized.

	Parameters:
		c - a control associated variable name to operate on
		a - (optional) one or more of the anchors: 'x', 'y', 'w' (width) and 'h' (height),
			optionally followed by a relative factor, e.g. "x h0.5"
		r - (optional) true to redraw controls, recommended for GroupBox and Button types

	Examples:
> "xy" ; bounds a control to the bottom-left edge of the window
> "w0.5" ; any change in the width of the window will resize the width of the control on a 2:1 ratio
> "h" ; similar to above but directrly proportional to height

	Remarks:
		To assume the current window size for the new bounds of a control (i.e. resetting)
		simply omit the second and third parameters.
		However if the control had been created with DllCall() and has its
		own parent window, the container AutoHotkey created GUI must be made default
		with the +LastFound option prior to the call.
		For a complete example see anchor-example.ahk.

	License:
		- Version 3.5.1 <http://www.autohotkey.net/~Titan/#anchor>
		- Simplified BSD License <http://www.autohotkey.net/~Titan/license.txt>
	 */
  static d
  GuiControlGet, p, Pos, %c%
  If !A_Gui or ErrorLevel
	Return
  i = x.w.y.h./.7.%A_GuiWidth%.%A_GuiHeight%.`n%A_Gui%:%c%=
  StringSplit, i, i, .
  d .= (n := !InStr(d, i9)) ? i9 :
  Loop, 4
	x := A_Index, j := i%x%, i6 += x = 3
	, k := !RegExMatch(a, j . "([\d.]+)", v) + (v1 ? v1 : 0)
	, e := p%j% - i%i6% * k, d .= n ? e . i5 : ""
	, RegExMatch(d, RegExReplace(i9, "([[\\\^\$\.\|\?\*\+\(\)])", "\$1")
	. "(?:([\d.\-]+)/){" . x . "}", v)
	, l .= InStr(a, j) ? j . v1 + i%i6% * k : ""
  r := r ? "Draw" :
  GuiControl, Move%r%, %c%, %l%
}
addfiles(type){
	global
	if (type = "drop")
	{
		Sort, a_guievent, N
		Loop, Parse, a_guievent, `n,`r
		{
			SplitPath, a_loopfield,file_name, file_dir
			StringLower, file_name, file_name
			LV_Add("",file_name,1)
		}
		LV_Modify(1, "Select")
		LV_GetText(cur_file, 1)
		if (inStr(cur_file, "jpeg"))
			ext := substr(cur_file, -3)
		else ext := substr(cur_file, -2)

		if (LockDept){
			GuiControl,, Preview, % "Preview: " . dept . "  - XX-XX-XX - Name." . ext
		}
		else GuiControl,, Preview, % "Preview: XX - XX-XX-XX - Name." . ext

		start()
	}
	if (type = "continue")
	{
		LV_GetText(cur_file, current)
		if (inStr(cur_file, "jpeg"))
			ext := substr(cur_file, -3)
		else
			ext := substr(cur_file, -2)
		LV_Modify(0, "-Select")
		LV_Modify(current-1, "", new_name)
		LV_Modify(current, "Select")
		if (!cur_file){
			Msgbox,68, % "End of image list", % "This was the last picture.`nDo you want to move the files to their corresponding directories?"
			IfMsgbox, No
				ExitApp
			moveFiles()
		}
		reset()
		start()
	}
}
start(){
	global
	WinActivate, % "CRename"
	GuiControl,, pic, % file_dir . "\" . cur_file

	GuiControl,Focus, dept
	Send, {Home}+{End}
	char :=
}
reset(){
	global

	GuiControl,,dept, % dept
	GuiControl,,day, % "day"
	GuiControl,,month, % "month"
	GuiControl,,year, % year
	GuiControl,,name, % "Name"

							if (LockDept){
		GuiControl,, Preview, % "Preview: " . dept . "  - XX-XX-XX - Name." . ext
	}
	else GuiControl,, Preview, % "Preview: XX - XX-XX-XX - Name." . ext

	char :=
}
gdisable(mode){
	global
	nlist := olddepdate . "|preview|SAC"
	Loop, Parse, nlist, |
	{
		if (LockDept && a_loopfield = "dept"){
			continue
		}
		GuiControl,%mode%,%a_loopfield%
	}
	if (mode = "disable")
	{
		GuiControl,,press, % "{ - } to break cycle"
		; change color
	}
	if (mode = "enable")
		GuiControl,,press, % "{ * }"

	start()
	char :=
}
crop(file){
	global iview
	Run, %iview% "%file% /info=%a_temp%\tmp /killmesoftly",,UseErrorLevel
	Sleep, 100
	FileRead, finfo, %a_temp%\tmp
	FileDelete, %a_temp%\tmp
	if (inStr(file, "jpeg"))
			ext := substr(file, -3)
		else
			ext := substr(file, -2)
	RegexMatch(finfo, "Image dimensions = ((\d+)\sx\s(\d+))", wh) ; Width = wh2 Height = wh3

	top := "(0,0," . wh2 . "," . RegexReplace(wh3/2,"\.\d+","") . ")"
	bottom := "(0," . RegexReplace(wh2/2,"\.\d+","") . "," . wh2 . "," . RegexReplace(wh3/2,"\.\d+","") . ")"

	RunWait, %iview% "%file% /crop=%top% /convert=%a_temp%\top.%ext% /killmesoftly",,UseErrorLevel
	Run, %iview% "%file% /crop=%bottom% /convert=%a_temp%\bottom.%ext% /killmesoftly",,UseErrorLevel
	FileDelete, res\*.ini
	return ext
}
moveFiles()
{
	Loop, Files, %A_WorkingDir%\*.*
	{
		OutputDebug, % A_LoopFileFullPath
	}
}
WM_LBUTTONUP(wParam, lParam)
{
	if (a_guicontrol = "dept" || a_guicontrol = "day" || a_guicontrol = "month"
	||  a_guicontrol = "year" || a_guicontrol = "name")
	{
		GuiControl,Focus,%a_guicontrol%
		Send, ^a
	}
}
;-

;+--> ; ---------[Hotkeys/Hotstrings]---------
#IfWinActive CRename
!Esc::ExitApp
Pause::Reload
F12::Pause

NumpadMult::Gosub, Previous

; Tab::
; Gosub, preview
; send {tab}
; return

; +Tab::
; Gosub, preview
; send +{tab}
; return

-::
NumpadSub::
 previous := 1
 gdisable("enable")
return

NumpadAdd::
 if (!LV_GetCount()){
	MsgBox,64,% "", % "There is no image to preview"
	return
 }
 toggle++
 if toggle = 1
 {
	t_ext := crop(file_dir . "\" . cur_file) ; save temporal extension for later use
	Gui, 2: +ToolWindow
	Gui, 2: add, Picture,w%wa_Right% h-1 x0 y0 +Border vpic2, % a_temp . "\top." . t_ext
	Gui, 2: Show
 }
 else if toggle = 2
	GuiControl, 2:, pic2, % a_temp . "\bottom." . t_ext
 else
 {
	Gui, 2: Destroy
	toggle :=
 }
return
#IfWinActive
;-

;+--> ; ---------[Includes]---------
;-

/*
 * =============================================================================================== *
 *                                            END OF FILE
 * =============================================================================================== *
 */
