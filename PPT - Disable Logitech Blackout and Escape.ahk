; script to mute logitech blackout button

; also adds notes scroller for operator using + and - keys on numeric keypad
; also prevents the start/stop button on logitech remote from quitting the slide deck (always issues start or F5 key)

; logitech clicker keys:
; [ < ] previous slide (normal behaviour)
; [ > ] next slide (normal behaviour)
; [ = ] ignored when in slide show  (this is the button below the [ < ])
; [ o ] ignored when in slide show  (this is the button below the [ > ])


; operator helper keys:
; F1 stop the presentation (works in both editor and slideshow)
; F2 start the presentation (or jump to it if in the editor)
; F3 start the presentation on the current slide (if slideshow is already running, it won't change current slide, just jumps to it)
; Numpad [-] Scroll the notes view up
; Numpad [+] Scroll the notes view down

; the F1 key was chosen for stop as it is the closest to the escape key, since escape is ignored

#Requires AutoHotkey v2.0

#SingleInstance Force ; if someone starts the script twice, just replace it in memory

#HotIf WinActive('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinActive('ahk_exe POWERPNT.EXE ahk_class screenClass')
; keys for when a powerpoint slideshow is in foreground
Esc::F5 ; Fix issue with logitech start/stop - always start
F1::Esc ; operator helper key - use F1 instead of escape (also mutes the F1 global help window which is "unhelpful")
vkBE:: {
    return ; fix issue with logitech blackout key - ignore it
}

NumpadSub::^Up   ; operator helper - numpad Minus = scroll notes up
NumpadAdd::^Down ; operator helper - numpad Plus  = scroll notes down

PgUp:: {
    Send "{PgUp}"
}

PgDn:: {
    ; if the operator has tabbed to the powerpoint editor, and the presenter clicks next, jump back to the slideshow and go next
    Send "{PgDn}"
}

+^Q:: {
    ; stop using this script
    ExitApp 0
}

#HotIf


#HotIf WinActive('ahk_exe POWERPNT.EXE ahk_class PPTFrameClass')
; keys for when the powerpoint editor is in foreground
F1:: {

    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
        Send "{Escape}"
    }

}

F2::F5
F3::+F5


F5:: {
    ; operator helper key - F5 starts the presentation, or jumps to it if it is in the background
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
    } else {
        Send "{F5}"
    }
}

Esc:: {

    ; when in the editor, the start/stop button (or escape/f5) will always force the slide show to the foreground
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
   ; } else {
      ;  Send "{F5}"
    }
}

+^Q:: {
    MsgBox "exiting PPT - Mute Logi Black"
    ; stop using this script
    ExitApp 0
}

#HotIf

; keys for every other app
+^Q:: {
    MsgBox "exiting PPT - Mute Logi Black"
    ; stop using this script
    ExitApp 0
}


PgUp:: {
    ; if the operator has tabbed to the powerpoint editor, and the presenter clicks previous, jump back to the slideshow and go previous
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
        SetPos()
    }
    Send "{PgUp}"
}

PgDn:: {
    ; if the operator has tabbed to the powerpoint editor, and the presenter clicks next, jump back to the slideshow and go next
    if WinExist('ahk_exe POWERPNT.EXE ahk_class PodiumParent') or WinExist('ahk_exe POWERPNT.EXE ahk_class screenClass') {
        WinActivate ; Use the window found by WinExist.
        SetPos()
    }
    Send "{PgDn}"
}

SetPos() {
    WinGetPos &X, &Y, &W, &H, "A"
    CoordMode "Mouse"
    MouseMove  (X + W - 20), (Y + H-20)  , 0
    Click 
 }