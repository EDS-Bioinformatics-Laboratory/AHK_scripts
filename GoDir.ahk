#Requires AutoHotkey v2
#SingleInstance Force

;F3+D will open file explorer in \AMC\Documents
F3 & d:: {
	target := "C:\Users\P075934\_data\Work\AMC\Documents"
    if FileExist(target) {
        Run('explorer.exe "' . target . '"')
    } else {
        ToolTip("Directory not found:`n" . target)
        Sleep 1500
        ToolTip()
    }
    return
}

;F3+p will open file explorer in \Projects\2025\Immunology
F3 & p:: {
	target := "C:\Users\P075934\_data\Work\AMC\Projects\2025\Immunology"
    if FileExist(target) {
        Run('explorer.exe "' . target . '"')
    } else {
        ToolTip("Directory not found:`n" . target)
        Sleep 1500
        ToolTip()
    }
    return
}

;F3+F will open explorer in FSS Projects
F3 & f:: {
	target := "C:\Users\P075934\_data\Dropbox\BioLab\FSS Projects"
    if FileExist(target) {
        Run('explorer.exe "' . target . '"')
    } else {
        ToolTip("Directory not found:`n" . target)
        Sleep 1500
        ToolTip()
    }
    return
}
