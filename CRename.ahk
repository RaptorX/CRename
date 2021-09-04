/*
	* =============================================================================================== *
	* Author           : RaptorX   <graptorx@gmail.com>
	* Script Name      : CRename
	* Script Version   : 1.9
	* Homepage         : -
	*
	* Creation Date    : September 18, 2010
	* Modification Date: May 06, 2021
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
	* =============================================================================================== *
 */

;+--> ; ---------[Directives]---------
#NoEnv
#SingleInstance Force

;+--> ; ---------[Basic Info]---------
s_name      := "CRename"                ; Script Name
s_version   := "1.9"                    ; Script Version
s_author    := "RaptorX"                ; Script Author
s_email     := "graptorx@gmail.com"     ; Author's contact email
;-

;+--> ; ---------[General Variables]---------
SysGet, screen_, MonitorWorkArea            ; Get the working area of the current screen
; --

;+--> ; ---------[User Configuration]---------
dataPath := FileExist(A_AppData) ? A_AppData "\" s_name "\" : ""
iviewPath := A_ScriptDir "\i_view32.exe"
emptyPath := A_ScriptDir "\empty.png"
depSettings := A_ScriptDir "\settings.ini"

memo =
(Ltrim
	Line1
	Line2
	Line3
	Line4
	Line5
)
;-

;+--> ; ---------[Main]---------
onMessage(0x202, "WM_LBUTTONUP")

if !FileExist(depSettings)
{
	Loop, 13
	{
		index := A_Index = 13 ? 99 : format("{:02}", A_Index)
		
		deps .= index "=`n"
	}
	
	FileAppend, % "[departments]`n" deps, % depSettings
}

for i,file in [iviewPath, emptyPath]
{
	if !FileExist(file)
	{
		if !FileExist(A_AppData "\" s_name)
			FileCreateDir, % A_AppData "\" s_name
		
		FileInstall, i_view32.exe, % iviewPath, % true
		FileInstall, empty.png, % emptyPath, % true
		break
	}
}

hpic := screen_Bottom - 50 - 10
lvsz := hpic - 280 +1

IniRead, year, % depSettings, % "restore", % "year", % SubStr(A_YYYY, -1)
IniRead, MoveFiles, % depSettings, % "restore", % "MoveFiles", % false
Gui, main:new, +MaximizeBox +MinSize hwnd$Main ;+Resize

Gui, add, Picture, w-1 h%hpic% y10 +Border vpic, % emptyPath
Gui, add, Text, w230 r5 x+10 cBlue vMemo, % memo
Gui, add, Edit, w40 y+10 cRed Section Limit2 Number vdept gPreview, % 00
Gui, add, Checkbox, w180 x+10 yp+3 right -TabStop checked%MoveFiles% vMoveFiles gSave, % "Move Files Automatically"
Gui, add, Edit, w40 xs y+10 cRed Limit2 Number vday gPreview, % "day"
Gui, add, Edit, w40 x+5 cRed Limit2 Number vmonth gPreview, % "month"
Gui, add, Edit, % "w40 x+5 " (!RegExMatch(year, "\d+") ? "cRed" : "") " Limit2 Number vyear gPreview", % year
Gui, add, Edit, w230 xs y+10 -wantreturn vname gPreview, % "Name"
Gui, add, Text, w200 r2 y+10 cRed center wrap vnamePreview, % "Preview: Invalid"
Gui, add, Text, w200 y+10 cBlue center vlbMagnify, % "{ + } to Magnify"
Gui, add, Button, w100 h25 y+10 disabled vnextPage gNext, % "Same && Next Page"
Gui, add, Button, w100 h25 x+10 +default vsaveContinue gSave, % "Save && Continue"
Gui, add, Text, w100 xs cBlue center disabled vpressAsterisk,% "{ * }"
Gui, add, Text, w100 x+10 cBlue center vpressEnter,% "{ Enter }"
Gui, add, ListView, w230 h%lvsz% xs y+15 grid sort altsubmit -hdr -tabstop vFileList glvHandler, % "fileName|filePath"

LV_ModifyCol(1, 225), LV_ModifyCol(2, 0)
Gui, Show,, % "CRename"

GuiControl, Focus, % "dept"
Send, {Home}+{End}
return

WM_LBUTTONUP(wParam, lParam, msg, hwnd)
{
	static EM_SETSEL := 0x00B1

	if (RegexMatch(A_GuiControl, "dept|day|month|year"))
		SendMessage, % EM_SETSEL, 0, -1,, % "ahk_id " hwnd
}

Preview(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global dept,day,month,year,name,namePreview,$Main
	isValidName := false
	isValidControl := false

	Gui, main:submit, nohide

	LV_GetText(filePath,row := (next:=LV_GetNext()) ? next : 1, 2)
	SplitPath, filePath,,, fileExt

	; independently verify each control
	if (A_GuiControl = "dept"  && (RegExMatch(dept, "(0[1-9]|1[0-2]|99)") || dept == ""))
	or (A_GuiControl = "day"   && RegExMatch(day  , "(0[0-9]|[1-2][0-9]|3[0-1])"))
	or (A_GuiControl = "month" && RegExMatch(month, "(0[0-9]|1[0-2])"))
	or (A_GuiControl = "year"  && RegExMatch(year , "\d{2}"))
	or (A_GuiControl = "name"  && !RegExMatch(name, "[\\/:*?""<>|]"))
			isValidControl := true

	if (!RegExMatch(A_GuiControl, "i)filelist|pic"))
	{
		Gui, font, % isValidControl ? "" : "cRed"
		GuiControl, font, % A_GuiControl
	}

	GuiControlGet, namePreview

	if  (name)
	and (RegExMatch(dept, "(0[1-9]|1[0-2]|99)") || dept == "")
	and RegExMatch(day  , "(0[0-9]|[1-2][0-9]|3[0-1])")
	and RegExMatch(month, "(0[0-9]|1[0-2])")
	and RegExMatch(year , "\d{2}")
	and !RegExMatch(name, "i)^name$")
	and !RegExMatch(name, "[\\/:*?""<>|]")
	and !RegExMatch(dept day month year, "i)dept|day|month|year")
		isValidName := true

	if  (A_GuiControl != "name" && isValidControl)
	and (WinActive("ahk_id " $Main))
	{
		if (A_GuiControl == "month" && year ~= "\d+")
			Send, {Tab 2}
		else
			Send, {Tab}
	}

	Gui, font, % isValidName ? "" : "cRed"
	GuiControl, font, % "namePreview"

	name := trim(RegExReplace(name, "\s+", " "))
	name := trim(RegExReplace(name, "((?:[[:upper:]]+)?[[:lower:]]+(?:[[:upper:]]+)?(?:[[:lower:]]+)?)", "$T{1}"))
	preview := (dept ? dept " - " : "") day "-" month "-" year " - " name (fileExt ? "."  fileExt : "")

	GuiControl,, % "namePreview", % "Preview: " (isValidName ? preview : "Invalid")
}

Save(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global depSettings

	Gui, main:submit, nohide

	if (A_GuiControl == "MoveFiles")
	{
		GuiControlGet, MoveFiles
		IniWrite, % MoveFiles, % depSettings, % "restore", % "MoveFiles"
		return
	}
	
	GuiControlGet, namePreview
	if (InStr(namePreview,"invalid") || !row := LV_GetNext(0, "F"))
	{
		MsgBox, % 0x10
		      , % "Error"
		      , % "The name you are trying to save is not valid"
		return
	}

	GuiControlGet, year
	GuiControlGet, MoveFiles

	IniWrite, % year, % depSettings, % "restore", % "year"
	IniWrite, % moveFiles, % depSettings, % "restore", % "MoveFiles"

	GuiControl, enable, % "nextPage"
	GuiControl, enable, % "pressAsterisk"

	LV_GetText(filePath, row ? row : 1, 2)

	SplitPath, filePath, fileName, fileDir, fileExt
	SetWorkingDir, % fileDir

	newFileName := RegExReplace(namePreview, "Preview: ")

	Try
		FileMove, % filePath, % newFileName
	Catch, fmError
	{
		MsgBox, % 0x10
		      , % "Error"
		      , % "There was an error saving " newFileName ".`n`n"
		      .   "The name you entered might already exist or you dont have permissions to write to that location.`n`n"
		      .   "Error Code:" A_LastError
		return
	}

	LV_Modify(row, "-select -focus", newFileName, A_WorkingDir "\" newFileName), LV_Modify(row+1, "focus select")
	LV_GetText(picPath, row+1)

	GuiControl,, % "pic", % A_WorkingDir "\" picPath	
	GuiControl,, % "day", % "day"
	GuiControl,, % "name", % "Name"
	GuiControl,, % "month", % "month"

	GuiControl, focus, % "dept"
	Send ^a

	if (!LV_GetNext(0, "F")) {
		finish()
	}
}

Next(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global
	static currentPage := 1
	static ctrlList := "dept|day|month|year|name|namePreview|saveContinue|pressEnter|lbDept|hkTab|lbDate|lbMagnify"

	Gui, main:default
	if (GuiEvent = "endCycle") {
		for i,ctrl in StrSplit(ctrlList, "|") {
			GuiControl, enable, % ctrl
		}
		
		GuiControl, focus, % "day"
		GuiControl, , % "pressAsterisk", % "{ * }"
		currentPage := 1
		return
	} else {
		for i,ctrl in StrSplit(ctrlList, "|") {
			GuiControl, disable, % ctrl
		}

		GuiControl, , % "pressAsterisk", % "{ * } or { - } to stop"
	}

	LV_GetText(currFile, row := LV_GetNext(0, "F"), 2)
	LV_GetText(prevFile, row - 1, 2)

	SplitPath, currFile,,, currFileExt
	SplitPath, prevFile,, prevFileDir, prevFileExt, prevFileName
	SetWorkingDir, % prevFileDir

	if (!RegExMatch(prevFile, "\bP\d{2}\b")) {
		Try
			FileMove, % prevFile, % newFileName := prevFileName " P" Format("{:02}",currentPage++) "." prevFileExt
		Catch, fmError
		{
			MsgBox, % 0x10
			, % "Error"
			, % "There was an error saving " newFileName ".`n`n"
			.   "The name you entered might already exist or you dont have permissions to write to that location.`n`n"
			.   "Error Code:" A_LastError
			return
		}

		LV_Modify(row - 1,"", newFileName, A_WorkingDir "\" newFileName)
	} else {
		RegExMatch(prevFileName, "\bP(\d{2})\b", match)
		currentPage := match1 + 1
		prevFileName := RegExReplace(prevFileName, "\b\sP\d{2}\b")
	}

	Try
		FileMove, % currFile, % newFileName := prevFileName " P" Format("{:02}",currentPage) "." currFileExt
	Catch, fmError
	{
		MsgBox, % 0x10
		, % "Error"
		, % "There was an error saving " newFileName ".`n`n"
		.   "The name you entered might already exist or you dont have permissions to write to that location.`n`n"
		.   "Error Code:" A_LastError
		return
	}

	LV_Modify(row,"-select -focus", newFileName, A_WorkingDir "\" newFileName)
	LV_Modify(row+1, "focus select"), LV_GetText(picPath, row+1)
	GuiControl,, % "pic", % A_WorkingDir "\" picPath

	if (!LV_GetNext(0, "F")) {
		finish()
	}
}

lvHandler(CtrlHwnd, GuiEvent, EventInfo, ErrLevel:="")
{
	global pic

	if (EventInfo && GuiEvent = "normal") {
		LV_GetText(picPath, EventInfo, 2)
		GuiControl,, pic, % picPath
	}
}

mainGuiDropFiles(GuiHwnd, FileArray, CtrlHwnd, X, Y)
{
	global $Main
	Gui, main:default

	for i,filePath in FileArray {
		SplitPath, filePath, fileName
		LV_Add("", fileName, filePath)
	}

	LV_Modify(1, "focus select")
	Preview(CtrlHwnd, "DropFiles", GuiHwnd)

	LV_GetText(picPath, 1, 2)
	GuiControl,, pic, % picPath

	WinActivate, % "ahk_id " $Main
}

mainGuiClose()
{
	ExitApp, 0
}

finish()
{
	global depSettings

	GuiControlGet, MoveFiles

	if (!MoveFiles)
		ExitApp, 0

	Loop, 13
	{
		index := A_Index = 13 ? 99 : A_Index
		IniRead, depPath, % depSettings, % "Departments", % depIndex := format("{:02}", index)
		if (!FileExist(depPath))
			continue

		Loop, Files, % A_WorkingDir "\*.*"
		{
			if (RegExMatch(A_LoopFileName, "^" depIndex "\s-\s"))
			{
				try
					FileMove, % A_LoopFileFullPath, % depPath, false
				Catch, err
				{
					overwriteLog .= A_LoopFileFullPath "`n"
					FileMove, % A_LoopFileFullPath, % "*-" A_MSec "-Dupe.*", false
			}
		}
	}
	}

	ExitApp, 0
}

crop(file)
{
	global iviewPath
	RunWait, %iviewPath% "%file% /info=%A_Temp%\tmp /killmesoftly"
	FileRead, finfo, % A_Temp "\tmp"
	FileDelete, % A_Temp "\tmp"
	
	RegExMatch(file, "\..*$", ext)
	RegexMatch(finfo, "O)Image dimensions = (?:(?<width>\d+)\sx\s(?<height>\d+))", size) ; Width = wh2 Height = wh3

	top := "(0,0," size.width "," Format("{:d}", size.height/2) ")"
	bottom := "(0," Format("{:d}", size.height/2) "," size.width "," Format("{:d}", size.height/2) ")"

	RunWait, %iviewPath% "%file% /crop=%top% /convert=%A_Temp%\top.%ext% /killmesoftly"
	RunWait, %iviewPath% "%file% /crop=%bottom% /convert=%A_Temp%\bottom.%ext% /killmesoftly"
	return ext
}

#IfWinActive, CRename

-::Next(0, "endCycle", 0)
NumpadSub::Next(0, "endCycle", 0)
NumpadMult::Next(0, "beginCycle", 0)
NumpadAdd::
	Gui, main:default
	
	if (!row := LV_GetNext(0, "F")){
		MsgBox,% 0x10, % "Error", % "Select a file to preview"
		return
	}
	
	LV_GetText(filePath, row, 2)
	SplitPath, filePath, fileName, fileDir

	toggle++
	if (toggle == 1) {
		ext := crop(filePath) ; save temporal extension for later use
		
		Gui, cropped:new, +toolwindow
		Gui, add, Picture,% "w" var:=(screen_Right /1.1) " x0 y0 +border vpic2", % A_Temp "\top." ext
		Gui, Show, noActivate
	} else if (toggle == 2) {
		GuiControl, cropped:, pic2, % A_Temp "\bottom." ext
	} else {
		for i,file in ["top", "bottom"]
			FileDelete, % A_Temp "\" file "." ext
		
		Gui, cropped:Destroy
		toggle := 0
	}
return

#IfWinActive