#Requires AutoHotkey v2.0
#SingleInstance Force


^m:: {
    wintitle_list := ['ahk_class OpusApp', 'ahk_class PPTFrameClass']
    wtitle_obj_popup := ['ahk_class bosa_sdm_msword', 'ahk_class NUIDialog']
    shortcut_key := ['o', 't']
    if WinActive(wintitle_list[1]) {
        state := 1
    } else if WinActive(wintitle_list[2]) {
        state := 2
    } else {
        MsgBox('Word or PowerPoint window is not active', 'Focus Required', 48)
        return
    }

    SendEvent '!njj'
    if !WinWait(wintitle_list[state],, 3) {
        MsgBox('Insert Object failed to open', 'Error', 16)
        return
    }
    sleep 100
    SendEvent('!' shortcut_key[state])
    sleep 100
    SendEvent('mathtype')
    sleep 100
    SendEvent '{Enter}'

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
