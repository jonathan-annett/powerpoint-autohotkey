#Persistent
#NoEnv
SetBatchLines, -1
CoordMode, Mouse, Screen

global regionStage := 0
global region := {}
global magnifierVisible := false

F8::
if GetKeyState("Shift", "P") {
    Gosub, ResetMagnifier
    Return
}

if (regionStage = 0) {
    MouseGetPos, region.x1, region.y1
    regionStage := 1
    Tooltip, Top-left set at %region.x1%, %region.y1%
    SetTimer, HideTooltip, -1000
}
else if (regionStage = 1) {
    MouseGetPos, region.x2, region.y2
    region.w := Abs(region.x2 - region.x1)
    region.h := Abs(region.y2 - region.y1)
    region.x := Min(region.x1, region.x2)
    region.y := Min(region.y1, region.y2)
    regionStage := 2
    Gosub, InitMagnifier
}
else if (regionStage = 2) {
    magnifierVisible := !magnifierVisible
    Gui, % (magnifierVisible ? "Show" : "Hide")
}
Return

+F8:: ; Shift+F8
ResetMagnifier:
regionStage := 0
magnifierVisible := false
Gui, Destroy
SetTimer, UpdateMag, Off
Tooltip, Region cleared
SetTimer, HideTooltip, -1000
Return

HideTooltip:
Tooltip
Return

InitMagnifier:
zoom := 2
destW := region.w * zoom
destH := region.h * zoom

Gui, +AlwaysOnTop -Caption +ToolWindow +E0x80000  ; WS_EX_LAYERED
Gui, Color, Black
Gui, Add, Picture, x0 y0 w%destW% h%destH% vMagPic
Gui, Show, w%destW% h%destH%, Magnifier
magnifierVisible := true
SetTimer, UpdateMag, 100
Return

UpdateMag:
    hDC := DllCall("GetDC", "Ptr", 0, "Ptr")
    hDest := DllCall("CreateCompatibleDC", "Ptr", hDC, "Ptr")
    hBmp := DllCall("CreateCompatibleBitmap", "Ptr", hDC, "Int", region.w, "Int", region.h, "Ptr")
    hOld := DllCall("SelectObject", "Ptr", hDest, "Ptr", hBmp, "Ptr")

    DllCall("BitBlt"
        , "Ptr", hDest
        , "Int", 0, "Int", 0
        , "Int", region.w, "Int", region.h
        , "Ptr", hDC
        , "Int", region.x, "Int", region.y
        , "UInt", 0x00CC0020)

    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hDC)

    GuiControl,, MagPic, HBITMAP:*%hBmp%

    DllCall("SelectObject", "Ptr", hDest, "Ptr", hOld)
    DllCall("DeleteDC", "Ptr", hDest)
    DllCall("DeleteObject", "Ptr", hBmp)
Return

GuiClose:
ExitApp
