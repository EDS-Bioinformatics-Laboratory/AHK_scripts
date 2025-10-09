#Requires AutoHotkey v2.0
#SingleInstance Force



^!f::  ;; ctrl-alt-f (format)
{
    if WinActive("ahk_exe WINWORD.EXE")
    {
        SendEvent("!hpg")  ;; open paragraph menu
        Sleep(10)
	    SendEvent("!gj")    ;; justify text
		Sleep(10)
		SendEvent '{Enter}'
		Sleep(10)
		SendEvent("!b")    ;; remove space before paragraph
		Sleep(10)
		SendEvent("0")
		Sleep(10)
		;;SendEvent '{Enter}'
		Sleep(10)
		SendEvent("!f")    ;; remove space after paragraph
		Sleep(10)
		SendEvent("0")
		Sleep(10)
		;;SendEvent '{Enter}'
		Sleep(10)
		SendEvent("!a")     ;; set line spacing
		Sleep(10)
        SendEvent("1.15")
		Sleep(10)
		SendEvent '{Enter}'
		Sleep(10)
		SendEvent("!hfca") ;; black font
		Sleep(10)
		SendEvent("!hff") ;; font	
        SendEvent("calibri")
		Sleep(10)
		SendEvent '{Enter}'
		Sleep(10)		
		SendEvent("!hfs") ;; font size
        SendEvent("11")
		Sleep(10)
		SendEvent '{Enter}'
	}
    else
    {
        MsgBox("Microsoft Word window not active.")
    }
}


