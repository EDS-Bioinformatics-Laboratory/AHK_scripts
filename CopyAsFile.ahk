#Requires AutoHotkey v2
#SingleInstance Force

; WIN+F copies filename text by simulating F2 + Ctrl+C
#f:: {
    ; Enter rename mode (F2)
    Send("{F2}")
    Sleep 150  ; wait for rename box to appear

    ; Copy the filename text
    Send("^c")
    Sleep 50   ; wait for clipboard

    ; Cancel rename (Escape key) so no changes are made
    Send("{Esc}")

    ; Optional: notify user
    ToolTip("Filename copied")
    Sleep 1000
    ToolTip()
    return
}
