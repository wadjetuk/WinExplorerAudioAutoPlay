#Requires AutoHotkey v2.0
#SingleInstance Force

; Auto-play selected audio in the ACTIVE File Explorer window (MCI backend, no COM/WMP).
; - Debounced, clipboard-safe, no beeps, context-menu aware
; - Tray menu toggle + hotkey to enable/disable Auto-Play at runtime
; - Robust filename handling: quotes, long names, multiple dots, 8.3 fallback,
;   and optional "copy to temp with simple name" fallback

; --- Config ---
AUDIO := Map(".mp3",1, ".m4a",1, ".wav",1, ".flac",1, ".ogg",1, ".oga",1, ".wma",1, ".aac",1, ".opus",1)
POLL_MS := 200               ; polling interval for auto-play
DEBOUNCE_MS := 800           ; ms before the same file can retrigger
ALIAS := "sel_audio"         ; MCI alias name
AUTO_PLAY := true            ; initial state: auto-play enabled
IDLE_GUARD_MS := 300         ; skip clipboard probe if keyboard used within N ms
COPY_FALLBACK := true        ; last-resort: copy to %TEMP% with a simple name if opens keep failing

; --- State ---
currentPath := ""
lastSeenPath := ""
lastPlayTick := 0
tempPlaying := ""            ; path of a temp file we created (if any)

; --- Hotkeys ---
^!p::PlayCurrentSelection()   ; Ctrl+Alt+P = play current selection once
^!s::StopPlayback()           ; Ctrl+Alt+S = stop playback
^!t::ToggleAutoPlay()         ; Ctrl+Alt+T = toggle Auto-Play on/off

; --- Tray menu ---
A_TrayMenu.Delete()  ; start clean
A_TrayMenu.Add("Auto-Play", ToggleAutoPlay)
A_TrayMenu.Add()
A_TrayMenu.Add("Play Current (Ctrl+Alt+P)", (*) => PlayCurrentSelection())
A_TrayMenu.Add("Stop (Ctrl+Alt+S)", (*) => StopPlayback())
A_TrayMenu.Add()
A_TrayMenu.Add("Exit", (*) => ExitApp())
if AUTO_PLAY
    A_TrayMenu.Check("Auto-Play")

; Start/stop polling
if AUTO_PLAY
    SetTimer(CheckAuto, POLL_MS)

; Handle MM_MCINOTIFY so we can close on completion and suppress default beep
OnMessage(0x3B9, EndNotify)  ; MM_MCINOTIFY

; ----------------------------------------------------------------
; Functions
; ----------------------------------------------------------------

ToggleAutoPlay(*) {
    global AUTO_PLAY, POLL_MS
    AUTO_PLAY := !AUTO_PLAY
    if AUTO_PLAY {
        SetTimer(CheckAuto, POLL_MS)
        A_TrayMenu.Check("Auto-Play")
    } else {
        SetTimer(CheckAuto, 0)
        A_TrayMenu.Uncheck("Auto-Play")
    }
}

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
    if IsBusy() || !IsExplorerActive()
        return

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
    ; Don’t interfere while you’re interacting or menus are open.
    global IDLE_GUARD_MS
    if WinExist("ahk_class #32768")                         ; context menu open
        return true
    if GetKeyState("RButton","P") || GetKeyState("LButton","P")
        return true
    if GetKeyState("Ctrl","P") || GetKeyState("Shift","P") || GetKeyState("Alt","P")
        return true
    if (A_TimeIdleKeyboard < IDLE_GUARD_MS)                 ; recent typing/shortcuts
        return true
    return false
}

GetSelectedPath_NoCom() {
    ; Clipboard-safe probe of Explorer selection.
    if !IsExplorerActive()
        return ""
    clipSaved := ClipboardAll()
    A_Clipboard := ""
    Send "^c"
    if !ClipWait(0.25) {
        A_Clipboard := clipSaved
        return ""
    }
    data := A_Clipboard
    A_Clipboard := clipSaved

    first := StrSplit(data, "`r`n")[1]
    if !first
        return ""
    ; Remove quotes if present, trim whitespace.
    if (SubStr(first,1,1) = '"' && SubStr(first,-1) = '"')
        first := SubStr(first, 2, StrLen(first)-2)
    first := Trim(first, " `t`r`n")
    if InStr(FileExist(first), "D")
        return ""  ; folder/virtual item
    return first
}

; --- Helpers for tricky filenames ---
CleanPath(p) {
    if !p
        return ""
    if (SubStr(p,1,1) = '"' && SubStr(p,-1) = '"')
        p := SubStr(p, 2, StrLen(p)-2)
    return Trim(p, " `t`r`n")
}

GetShortPath(p) {
    buf := Buffer(32768 * 2, 0) ; wide char buffer
    len := DllCall("kernel32\GetShortPathNameW", "WStr", p, "Ptr", buf.Ptr, "UInt", 32768, "UInt")
    return len ? StrGet(buf.Ptr, len, "UTF-16") : ""
}

MakeTempCopySimple(p) {
    ; Last-resort workaround: copy to %TEMP% with a simple name + same extension
    try {
        SplitPath p, , , &ext
        simp := A_Temp . "\mci_preview." . ext
        FileCopy p, simp, true
        return simp
    } catch {
        return ""
    }
}

StartPlayback(path) {
    global currentPath, ALIAS, tempPlaying, COPY_FALLBACK
    StopPlayback()  ; close any previous (also cleans tempPlaying)

    path := CleanPath(path)
    if !FileExist(path)
        return

    ok := false
    ; 1) Try plain open
    if MciOK(Format('open "{}" alias {}', path, ALIAS)) {
        ok := true
    } else {
        ; 2) Try type by extension
        SplitPath path, , , &ext
        ext := StrLower(ext)
        typeStr := ""
        if (ext = "wav")
            typeStr := "waveaudio"
        else if (ext = "mp3" || ext = "aac" || ext = "m4a" || ext = "wma" || ext = "flac" || ext = "ogg" || ext = "oga" || ext = "opus")
            typeStr := "mpegvideo"
        if (!ok && typeStr != "")
            ok := MciOK(Format('open "{}" type {} alias {}', path, typeStr, ALIAS))
        ; 3) Try 8.3 short path
        if !ok {
            sp := GetShortPath(path)
            if sp {
                ok := MciOK(Format('open "{}" alias {}', sp, ALIAS))
                if !ok && (typeStr != "")
                    ok := MciOK(Format('open "{}" type {} alias {}', sp, typeStr, ALIAS))
            }
        }
        ; 4) Optional: copy to TEMP with a simple name, then open
        if !ok && COPY_FALLBACK {
            simp := MakeTempCopySimple(path)
            if simp {
                tempPlaying := simp
                ok := MciOK(Format('open "{}" alias {}', simp, ALIAS))
                if !ok && (typeStr != "")
                    ok := MciOK(Format('open "{}" type {} alias {}', simp, typeStr, ALIAS))
                if !ok {
                    ; cleanup temp if even that failed
                    TryDeleteTemp()
                }
            }
        }
    }

    if !ok
        return

    ; Play once (no loop) and request notify so we can cleanly close on end
    if MciOK("play " . ALIAS . " notify")
        currentPath := path
    else
        StopPlayback()
}

StopPlayback() {
    global currentPath, ALIAS
    MciOK("stop " . ALIAS)
    MciOK("close " . ALIAS)
    TryDeleteTemp()
    currentPath := ""
}

TryDeleteTemp() {
    global tempPlaying
    if tempPlaying {
        try FileDelete tempPlaying
        tempPlaying := ""
    }
}

MciOK(cmd) {
    return DllCall("winmm\mciSendStringW", "WStr", cmd, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt") = 0
}

EndNotify(wParam, lParam, msg, hwnd) {
    ; Handle MM_MCINOTIFY to avoid OS chime and to clear state on completion.
    global ALIAS, currentPath
    if (wParam = 1) {  ; MCI_NOTIFY_SUCCESSFUL
        MciOK("close " . ALIAS)
        TryDeleteTemp()
        currentPath := ""
    }
    return true  ; swallow message to prevent default beep
}

Shutdown(*) {
    StopPlayback()
    ExitApp
}
