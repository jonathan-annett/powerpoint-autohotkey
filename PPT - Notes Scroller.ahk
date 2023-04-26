; script to replace logitech start/stop and blackout buttons with a notes scroller system
; also adds notes scroller for operator using + and - keys on numeric keypad

; logitech keys:
; [ < ] previous slide (normal behaviour)
; [ > ] next slide (normal behaviour)
; [ = ] scroll notes up (this is the the button below the [ < ])
; [ o ] scroll notes down (this is the the button below the [ > ])


; operator helper keys:
; F1 stop the presentation
; F2 start the presentation (or jump to it if in the editor)
; F3 start the presentation on the current slide (or jump to slideshow if it is already running - won't change current slide)
; Numpad [-] Scroll the notes view up
; Numpad [+] Scroll the notes view down

; more info
; the start/stop key in a logitech remote alternates between escape and f5.
;             the blackout key (to the right of the start/stop) sends the OEM_PERIOD or virtualKey 0xBE (the "." key)
; we trap these keys and replace them with ctrl-up (in the case of f5 or escape), and ctrl-down
; this is used by powerpoint to scrol the notes view

; so the operator has a way to stop and start the slide deck (besides using the mouse)
; we provide:  F2 to start the slideshow (or bring it to the foreground)
;              F1 to exit the slide show
; these are basically because the Escape and F5 keys have been used to provide the scrolling

#Requires AutoHotkey v2.0

#SingleInstance Force

#HotIf WinActive('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinActive('ahk_exe POWERPNT.EXE ahk_class screenClass')
; keys for when a powerpoint slideshow is in foreground
Esc::^Up ;
F5::^Up ;
F1::Esc ;
vkBE::^Down ; replace logtech blackout with notes scroll down

NumpadSub::^Up   ; operator helper - numpad Minus = scroll notes up
NumpadAdd::^Down ; operator helper - numpad Plus  = scroll notes down

+^Q:: {
    ; exit silently
    ExitApp 0
}


#HotIf


#HotIf WinActive('ahk_exe POWERPNT.EXE ahk_class PPTFrameClass')
; keys for when the powerpoint editor is in foreground
F1:: {
    ; operator helper key for when in editor mode - stop the presentation

    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
        Send "{Escape}"
    }

}

F2::F5
F3::+F5

+^Q:: {
    MsgBox "exiting PPT - Notes Scroller"
    ExitApp 0
}

#HotIf

+^Q:: {
    MsgBox "exiting PPT - Notes Scroller"
    ExitApp 0
}

PgUp:: {
    ; if presenter clicks back when operator is in editor, jump to slideshown and select previous slide
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
    }
    Send "{PgUp}"
}

PgDn:: {
    ; if presenter clicks next when operator is in editor, jump to slideshown and select next slide
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
    }
    Send "{PgDn}"
}

Esc:: {
    ; scroll notes up
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
        Send "^{Up}"
    }
}

F5:: {
    ; scroll notes down
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
        Send "^{Up}"
    }
}
