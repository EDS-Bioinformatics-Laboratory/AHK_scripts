; ==============================
; Project Metadata
; ==============================
ProjectName := "MathType Object Inserter"
Version     := "1.0.0"
Author      := "Christoffel"
DateCreated := "2025-10-01"
; ==============================

#Requires AutoHotkey v2.0
#SingleInstance Force

CoordMode 'ToolTip', 'Screen'
CoordMode 'Mouse', 'Client'
CoordMode 'Pixel', 'Client'
SendMode 'Event'

hotkeys := Map()
hotkeys.start       := LoadSetting('Hotkeys', 'start', '^m')
hotkeys.pause       := LoadSetting('Hotkeys', 'pause', 'Esc')
hotkeys.tog_tooltip := LoadSetting('Hotkeys', 'toggle_tooltip', '+F1')
hotkeys.reload      := LoadSetting('Hotkeys', 'reload', '+Esc')
hotkeys.exit        := LoadSetting('Hotkeys', 'exit', '^Esc')

hotkeys.pause       := '$' hotkeys.pause
Initialize

Main(var, ThisHotkey) {
    global
    try
        KeyWait ThisHotkey
    KeyWait 'Ctrl'
    KeyWait 'Shift'
    KeyWait 'Alt'
    KeyWait 'LWin'
    MainTooltip('Running.')
    Hotkey(hotkeys.start, 'Off')
    Hotkey(hotkeys.pause, 'On')
    wintitle_list := ['ahk_class OpusApp', 'ahk_class PPTFrameClass']
    wtitle_obj_popup := ['ahk_class bosa_sdm_msword', 'ahk_class NUIDialog']
    shortcut_key := ['o', 't']
    if WinActive(wintitle_list[1]) {
        state := 1
    } else if WinActive(wintitle_list[2]) {
        state := 2
    } else {
        MsgBox('Word or PowerPoint window is not active', 'Focus Required', 48)
        return ReadyState()
    }

    SendEvent '!njj'
    if !WinWait(wintitle_list[state],, 3) {
        MsgBox('Insert Object failed to open', 'Error', 16)
        return ReadyState()
    }
    sleep 100
    SendEvent('!' shortcut_key[state])
    sleep 100
    SendEvent('mathtype')
    sleep 100
    SendEvent '{Enter}'

    return ReadyState()
}

class SettingsGUI {
    static __New() {
        this.gui_obj := Gui('AlwaysOnTop')
        this.gui_obj.AddText('xm w100', 'Setting1')
        this.ctrl_setting1 := this.gui_obj.AddEdit('w100 yp', 'Setting1')
        this.gui_obj.AddText('xm w100', 'Setting2')
        this.ctrl_setting2 := this.gui_obj.AddEdit('w100 yp', 'Setting2')
        this.ctrl_btn_save := this.gui_obj.AddButton('w100', 'Save')
        this.ctrl_btn_cancel := this.gui_obj.AddButton('w100', 'Cancel')

        this.ctrl_btn_save.OnEvent('Click', this.Save.Bind(this))
        this.ctrl_btn_cancel.OnEvent('Click', this.Cancel.Bind(this))
    }

    static Save(*) {
        this.gui_obj.Submit()
    }
    
    static Cancel(*) {
        this.Hide()
    }

    static Show() {
        this.gui_obj.Show()
    }

    static Hide() {
        this.gui_obj.Hide()
    }
}

LoadSetting(sect, key, default_or_mode := 'NONE') {
    settings_file := 'settings.ini'
    val := IniRead(settings_file, sect, key, '')
    if RegExMatch(val, "^\{.*\}$") {
        val := Trim(val, '{}')
        _obj := {}
        for v in StrSplit(val, ',', ' `t') {
            _ := StrSplit(v, ':', ' `t')
            _obj.%_[1]% := _[2]
        }
        val := _obj
        return val
    } else if val
        return val

    switch default_or_mode {
        case 'NONE':
            val := ''
        case '1coord':
            while !GetKeyState('LControl', 'P') {
                MouseGetPos(&x, &y, &win)
                win := WinGetTitle(win) ' ahk_class ' WinGetClass(win)
                ToolTip('Press [Left Ctrl] to set coordinate for: ' key '`n(' x ',' y ' | ' win ')')
                sleep 20
            }
            KeyWait('LControl')
            ToolTip
            val := {x: x, y: y, win: win}
        case '2coord':
            val := {}
            while !GetKeyState('LControl', 'P') {
                MouseGetPos(&x, &y, &win)
                win := WinGetTitle(win) ' ahk_class ' WinGetClass(win)
                ToolTip('Press [Left Ctrl] to set TOP-LEFT coordinate for: ' key '`n(' x ',' y ' | ' win ')')
                sleep 20
            }
            val.x1 := x
            val.y1 := y
            val.win := win
            KeyWait('LControl')
            while !GetKeyState('LControl', 'P') {
                MouseGetPos(&x, &y, &win)
                win := WinGetTitle(win) ' ahk_class ' WinGetClass(win)
                ToolTip('Press [Left Ctrl] to set BOTTOM-RIGHT coordinate for: ' key '`n(' x ',' y ' | ' win ')')
                sleep 20
            }
            KeyWait('LControl')
            ToolTip
            val.x2 := x
            val.y2 := y
        default:
            val := default_or_mode 

    }
    if IsObject(val) {
        _ := '{'
        for k, v in val.OwnProps() {
            _ .= k ': ' v ', '
        }
        _ := Trim(_, ', ') '}'
        IniWrite(_, settings_file, sect, key)
    } else {
        IniWrite(val, settings_file, sect, key)
    }
    return val
}

Initialize() {
    global
    ; HotIfWinActive("")
    Hotkey(hotkeys.start, Main.Bind('Start'))
    ; HotIfWinActive
    Hotkey(hotkeys.tog_tooltip, ToggleTooltip)
    Hotkey(hotkeys.pause, PauseScript)
    Hotkey(hotkeys.reload, (*) => Reload())
    Hotkey(hotkeys.exit, (*) => ExitApp())
    
    tog_tooltip := True
    
    MainTooltip('Initializing... Please follow the instructions.')

    coords := {}

    ReadyState
}

SelectScreenRegion(Key:='LButton', Color := "Lime", Transparent:= 80) {
    ToolTip('SELECTION MODE', 5, 5, 19)
    hotkey('$' Key, (*) => 0, 'On')
    KeyWait(Key, 'D')
	CoordMode("Mouse", "Screen")
	MouseGetPos(&sX, &sY)
	ssrGui := Gui("+AlwaysOnTop -caption +Border +ToolWindow +LastFound -DPIScale")
	WinSetTransparent(Transparent)
	ssrGui.BackColor := Color
	Loop 
	{
		Sleep 10
		MouseGetPos(&eX, &eY)
		W := Abs(sX - eX), H := Abs(sY - eY)
		X := Min(sX, eX), Y := Min(sY, eY)
		ssrGui.Show("x" X " y" Y " w" W " h" H)
	} Until !GetKeyState(Key, "p")
	ssrGui.Destroy()
    hotkey('$' Key, 'Off')
    ToolTip(, , , 19)

	Return { X: X, Y: Y, W: W, H: H, X2: X + W, Y2: Y + H }
}

ReadyState() {
    MainTooltip('Ready (' ProjectName ' v' Version ')')
    ; HotIfWinActive("")
    Hotkey(hotkeys.start, "On")
    Hotkey(hotkeys.pause, "Off")
    ; HotIfWinActive
    Hotkey(hotkeys.reload, "On")
    Hotkey(hotkeys.exit, "On")
}

MainTooltip(var) {
    str_hotkeys := '    ' FormatHotkey(hotkeys.start)         '`t: Start'
    ; for k, v in hotkeys.OwnProps() {
    ;     str_hotkeys .= '    ' FormatHotkey(v) '`t`t: ' k '`n'
    ; }
    ; str_hotkeys := Trim(str_hotkeys, '`n')
    if tog_tooltip {
        ToolTip(
        (
            var '`nHotkeys:`n' str_hotkeys '
                ' FormatHotkey(hotkeys.pause)         '`t`t: Pause/Resume
                ' FormatHotkey(hotkeys.reload)        '`t: Stop/Reload
                ' FormatHotkey(hotkeys.tog_tooltip)   '`t: Show/Hide info
                ' FormatHotkey(hotkeys.exit)          '`t: Exit'
        ), 5, 5, 2)
        ToolTip('',,, 11)
        ToolTip('',,, 12)
    } else {
        ToolTip ,,, 2
        ToolTip ,,, 11
        ToolTip ,,, 12
    }
}

ToggleTooltip(ThisHotkey) {
    global tog_tooltip
    tog_tooltip := !tog_tooltip
    ReadyState()
}

FormatHotkey(hotkey) {
    hotkey := StrReplace(hotkey, "+", "Shift+")
    hotkey := StrReplace(hotkey, "^", "Ctrl+")
    hotkey := StrReplace(hotkey, "!", "Alt+")
    hotkey := StrReplace(hotkey, "#", "Win+")
    return hotkey
}

PauseScript(*) {
    Pause(-1)
    if A_IsPaused
        MainTooltip('Paused.')
    else
        MainTooltip('Running...')
}

RunAsAdmin() {
    full_command_line := DllCall("GetCommandLine", "str")

    if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
    {
        try
        {
            if A_IsCompiled
                Run '*RunAs "' A_ScriptFullPath '" /restart'
            else
                Run '*RunAs "' A_AhkPath '" /restart "' A_ScriptFullPath '"'
        }
        ExitApp
    }
}
