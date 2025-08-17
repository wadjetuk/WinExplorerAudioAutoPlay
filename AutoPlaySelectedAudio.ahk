#Requires AutoHotkey v2.0
#SingleInstance Force

; Auto-play selected audio in the ACTIVE File Explorer window (no COM/WMP).
; Uses MCI (winmm). Debounced. Plays once (notify). No chime on right-click.
; Now clipboard-safe: never copies while you're actively using Ctrl/Shift/Alt,
; the mouse, or the keyboard (idle guard).

; --- Config ---
AUDIO := Map(".mp3",1, ".m4a",1, ".wav",1, ".flac",1, ".ogg",1, ".oga",1, ".wma",1, ".aac",1, ".opus",1)
POLL_MS := 200             ; polling interval for auto-play
DEBOUNCE_MS := 800         ; ms before the same file can retrigger
ALIAS := "sel_audio"       ; MCI alias name
AUTO_PLAY := true          ; set false if you only want hotkey mode
IDLE_GUARD_MS := 300       ; don't probe clipboard if keyboard active in last N ms

; --- State ---
currentPath := ""
lastSeenPath := ""
lastPlayTick := 0

; --- Hotkeys ---
^!p::PlayCurrentSelection()   ; Ctrl+Alt+P = play once
^!s::StopPlayback()           ; Ctrl+Alt+S = stop

; Auto-play loop (enabled by AUTO_PLAY)
if AUTO_PLAY
    SetTimer(CheckAuto, POLL_MS)

; Handle MCI notify so we can close on completion and suppress default beep
OnMessage(0x3B9, EndNotify)  ; MM_MCINOTIFY

; ----------------------------------------------------------------
; Functions
; ----------------------------------------------------------------

PlayCurrentSelection() {
    global AUDIO, currentPath, lastSeenPath, lastPlayTick, DEBOUNCE_MS
    if IsBusy() || !IsExplorerActive()
        return
    path := GetSelectedPath_NoCom()
    if !path
        return
    SplitPath path, , , &ext
    ext := "." . StrLower(ext)
    if !AUDIO.Has(ext)
        return
    if (path = currentPath) || (A_TickCount - lastPlayTick < DEBOUNCE_MS)
        return
    lastPlayTick := A_TickCount
    StartPlayback(path)
    lastSeenPath := path
}

CheckAuto() {
    global AUDIO, currentPath, lastSeenPath, lastPlayTick, DEBOUNCE_MS
    if IsBusy() || !IsExplorerActive() {
        ; If we’re not in a good state to probe, don’t touch clipboard.
        return
    }

    path := GetSelectedPath_NoCom()
    if !path {
        if currentPath
            StopPlayback()
        lastSeenPath := ""
        return
    }

    if (path = lastSeenPath)
        return
    lastSeenPath := path

    SplitPath path, , , &ext
    ext := "." . StrLower(ext)
    if !AUDIO.Has(ext) {
        if currentPath
            StopPlayback()
        return
    }

    if (A_TickCount - lastPlayTick < DEBOUNCE_MS)
        return
    lastPlayTick := A_TickCount
    StartPlayback(path)
}

IsExplorerActive() {
    hwnd := WinActive("ahk_class CabinetWClass")
    if !hwnd
        hwnd := WinActive("ahk_class ExploreWClass")
    return !!hwnd
}

IsBusy() {
    ; Don’t interfere while user is interacting.
    global IDLE_GUARD_MS
    if WinExist("ahk_class #32768")                    ; context menu open
        return true
    if GetKeyState("RButton","P") || GetKeyState("LButton","P")
        return true
    if GetKeyState("Ctrl","P") || GetKeyState("Shift","P") || GetKeyState("Alt","P")
        return true
    if (A_TimeIdleKeyboard < IDLE_GUARD_MS)            ; recent typing/shortcuts
        return true
    return false
}

GetSelectedPath_NoCom() {
    ; Clipboard-safe probe: saves & restores clipboard immediately.
    if !IsExplorerActive()
        return ""
    clipSaved := ClipboardAll()
    A_Clipboard := ""          ; clear to detect our own copy result
    Send "^c"
    if !ClipWait(0.25) {
        A_Clipboard := clipSaved
        return ""
    }
    data := A_Clipboard
    A_Clipboard := clipSaved   ; restore exactly what user had

    first := StrSplit(data, "`r`n")[1]
    if !first
        return ""
    if (SubStr(first,1,1) = '"' && SubStr(first,-1) = '"')
        first := SubStr(first, 2, StrLen(first)-2)
    if InStr(FileExist(first), "D")
        return ""  ; folder or virtual item
    return first
}

StartPlayback(path) {
    global currentPath, ALIAS
    StopPlayback()  ; close any previous

    ok := MciOK(Format('open "{}" alias {}', path, ALIAS))
    if !ok {
        SplitPath path, , , &ext
        ext := StrLower(ext)
        typeStr := ""
        if (ext = "wav")
            typeStr := "waveaudio"
        else if (ext = "mp3" || ext = "aac" || ext = "m4a" || ext = "wma" || ext = "flac" || ext = "ogg" || ext = "oga" || ext = "opus")
            typeStr := "mpegvideo"
        if (typeStr != "")
            ok := MciOK(Format('open "{}" type {} alias {}', path, typeStr, ALIAS))
        if !ok
            return
    }

    if MciOK("play " . ALIAS . " notify")
        currentPath := path
}

StopPlayback() {
    global currentPath, ALIAS
    MciOK("stop " . ALIAS)
    MciOK("close " . ALIAS)
    currentPath := ""
}

MciOK(cmd) {
    return DllCall("winmm\mciSendStringW", "WStr", cmd, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt") = 0
}

EndNotify(wParam, lParam, msg, hwnd) {
    global ALIAS, currentPath
    if (wParam = 1) {  ; MCI_NOTIFY_SUCCESSFUL
        MciOK("close " . ALIAS)
        currentPath := ""
    }
    return true  ; swallow message to prevent default beep
}

Shutdown(*) {
    StopPlayback()
    ExitApp
}
