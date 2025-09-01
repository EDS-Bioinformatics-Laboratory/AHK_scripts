#Requires AutoHotkey v2.0
#SingleInstance Force

;; AutoHotKey script
;; https://www.autohotkey.com/

;; Description:
;;    Creates a new text file and opens it in Notepad++.

;; Developer
;;    Christoffel
;;    https://www.fiverr.com/christoffel16
;;    August 2025

;; License: GPL

CoordMode 'ToolTip', 'Screen'


global filename := ''

F1 & t:: { ;;text file
    target_dir := SetTargetDir()
    filename := EnterFileName(target_dir,"")
    output_path := target_dir '\' filename '.txt'
    if !FileExist(output_path) {
        FileAppend('', output_path)
        MainTooltip('File created: ' output_path)
    } else {
        MainTooltip('File already exists: ' output_path)
    }
    Run('notepad++.exe "' output_path '"')
}

F1 & m:: { ;; markdown file
    target_dir := SetTargetDir()
    filename := EnterFileName(target_dir,"")
    output_path := target_dir '\' filename '.md'
    if !FileExist(output_path) {
        FileAppend('', output_path)
        MainTooltip('File created: ' output_path)
    } else {
        MainTooltip('File already exists: ' output_path)
    }
    Run('Typora.exe "' output_path '"')
}

F1 & w:: {  ;;Word file
    target_dir := SetTargetDir()
    filename := EnterFileName(target_dir,"")
    output_path := target_dir '\' filename '.docx'
    if !FileExist(output_path) {
        FileAppend('', output_path)
        MainTooltip('File created: ' output_path)
    } else {
        MainTooltip('File already exists: ' output_path)
    }
    Run('WINWORD.exe "' output_path '"')
}

F1 & p:: { ;;PowerPoint file
    target_dir := SetTargetDir()
    filename := EnterFileName(target_dir,"")
    output_path := target_dir '\' filename '.pptx'
    if !FileExist(output_path) {
        FileAppend('', output_path)
        MainTooltip('File created: ' output_path)
    } else {
        MainTooltip('File already exists: ' output_path)
    }
    Run('POWERPNT.exe "' output_path '"')
}

F1 & e:: { ;;Excel file
    target_dir := SetTargetDir()
    filename := EnterFileName(target_dir,"")
    output_path := target_dir '\' filename '.xlsx'
    if !FileExist(output_path) {
        ;;FileAppend('', output_path)
        MainTooltip('File created: ' output_path)
    } else {
        MainTooltip('File already exists: ' output_path)
    }
    ;;Run('EXCEL.exe "' output_path '"')
    xl := ComObject("Excel.Application")
    xl.Visible := true
    wb := xl.Workbooks.Add() ; new workbook with one sheet
    wb.SaveAs(output_path, 51)
    ;;wb.Close(false)

}



GetActiveExplorerPath() {
    for window in ComObject("Shell.Application").Windows {
        try {
            if (WinActive("ahk_id " window.HWND)) {
                return window.Document.Folder.Self.Path
            }
        }
    }
    return ""
}

SetTargetDir() {
    if WinActive('ahk_class CabinetWClass ahk_exe explorer.exe')
        target_dir := GetActiveExplorerPath()
    else
        target_dir := A_Desktop

    return target_dir
}

EnterFileName(target_dir,filename) {
    loop {
        ib := InputBox('Input the name of the file to create on:`n' target_dir, 'Input filename', 'h110', filename)
        if ib.Result = 'Cancel'
            return MainTooltip('canceled by user')
        filename := ib.Value
        ; Check for valid filename: not empty, no invalid chars
        if filename != '' && !RegExMatch(filename, '[\\/:*?"<>|]')
            break
        MsgBox('Please enter a valid filename (no \ / : * ? " < > | and not empty).', 'Invalid Filename', 48)
    }
    return filename
}

MainTooltip(text) {
    ToolTip(text)
    SetTimer(() => ToolTip(), -2000)
    return
}
