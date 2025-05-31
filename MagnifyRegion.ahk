#Requires AutoHotkey v2.0
#SingleInstance Force
SetTimer(HideTooltip, -1)  ; Predeclare
global region := Map(), regionStage := 0, magnifierVisible := false
global zoom := 2, updateTimer := 0, myGui := unset, picCtrl := unset

F8:: {
    if GetKeyState("Shift") {
        ResetMagnifier()
        return
    }

    if regionStage = 0 {
        MouseGetPos &x, &y
        region["x1"] := x, region["y1"] := y
        regionStage := 1
        ToolTip("Top-left set at " region["x1"] ", " region["y1"])
        SetTimer(HideTooltip, -1000)
    }
    else if regionStage = 1 {
        MouseGetPos &x, &y
        region["x2"] := x, region["y2"] := y

        region["x"] := Min(region["x1"], region["x2"])
        region["y"] := Min(region["y1"], region["y2"])
        region["w"] := Abs(region["x2"] - region["x1"])
        region["h"] := Abs(region["y2"] - region["y1"])

        regionStage := 2
        InitMagnifier()
    }
    else if regionStage = 2 {
        magnifierVisible := !magnifierVisible
        if magnifierVisible
            myGui.Show()
        else
            myGui.Hide()
    }
}

+F8::ResetMagnifier()

ResetMagnifier() {
    global regionStage, magnifierVisible, myGui, updateTimer
    regionStage := 0
    magnifierVisible := false
    if IsSet(myGui)
        myGui.Destroy()
    ToolTip("Region cleared")
    SetTimer(HideTooltip, -1000)
    if updateTimer
        SetTimer(UpdateMag, 0)
}

HideTooltip(*) => ToolTip()

InitMagnifier() {
    global region, zoom, myGui, picCtrl, updateTimer, magnifierVisible
    destW := region["w"] * zoom
    destH := region["h"] * zoom

    myGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    myGui.BackColor := "Black"
    picCtrl := myGui.Add("Picture", "x0 y0 w" destW " h" destH " vMagPic")
    myGui.Show("w" destW " h" destH)
    magnifierVisible := true

    updateTimer := SetTimer(UpdateMag, 100)
}

UpdateMag(*) {
    global region, picCtrl

    hDC := DllCall("GetDC", "ptr", 0, "ptr")
    hDest := DllCall("CreateCompatibleDC", "ptr", hDC, "ptr")
    hBmp := DllCall("CreateCompatibleBitmap", "ptr", hDC, "int", region["w"], "int", region["h"], "ptr")
    hOld := DllCall("SelectObject", "ptr", hDest, "ptr", hBmp, "ptr")

    DllCall("BitBlt", "ptr", hDest, "int", 0, "int", 0, "int", region["w"], "int", region["h"],
        "ptr", hDC, "int", region["x"], "int", region["y"], "uint", 0x00CC0020)

    DllCall("ReleaseDC", "ptr", 0, "ptr", hDC)

    picCtrl.Value := "HBITMAP:*" hBmp

    DllCall("SelectObject", "ptr", hDest, "ptr", hOld)
    DllCall("DeleteDC", "ptr", hDest)
    DllCall("DeleteObject", "ptr", hBmp)
}
