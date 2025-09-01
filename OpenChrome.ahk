#Requires AutoHotkey v2
#SingleInstance Force

; === configure your pages here ===
pages1 := [
    "https://www.google.com/",
	"https://www.google.com/",
	"https://bioinformaticslaboratory.eu/",
	"https://bioinformaticslaboratory.eu/wp-admin",
	"https://pubmed.ncbi.nlm.nih.gov/"
]

pages2 := [
    "https://www.google.com/",
	"https://www.google.com/",
]

pages3 := [
    "https://chatgpt.com/",
    "https://www.google.com/"
]

; Path to Chrome — adjust if you installed in a non-standard location
chromePath := "C:\Program Files\Google\Chrome\Application\chrome.exe"

if !FileExist(chromePath) {
    MsgBox("Chrome executable not found:`n" . chromePath)
    ExitApp
}

F2 & W:: { ; F2+W 
    ; Build argument string: --new-window followed by all URLs
    args := "--new-window"
    for index, url in pages1
        args .= " " . url

    Run('"' . chromePath . '" ' . args)
    return
}

F2 & P:: { ; F2+P  
    ; Build argument string: --new-window followed by all URLs
    args := "--new-window"
    for index, url in pages2
        args .= " " . url

    Run('"' . chromePath . '" ' . args)
    return
}

F2 & G:: { ; F2+G 
    ; Build argument string: --new-window followed by all URLs
    args := "--new-window"
    for index, url in pages3
        args .= " " . url

    Run('"' . chromePath . '" ' . args)
    return
}



