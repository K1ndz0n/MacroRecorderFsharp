module Program

open System
open System.Collections.Generic
open System.Runtime.InteropServices
open System.Windows.Forms
open Microsoft.FSharp.Core
open System.Diagnostics
open System.Collections.ObjectModel
open Structures
open Classes
        
                
let macroBuf = ObservableCollection<MacroEvent>() // lista nagranych zdarzeń
let keyStrokeBuf = List<string>() // Bufor do tworzenia KeyStroke
let recordedMacros = ObservableCollection<Macro>()

let mutable isHotkeyStarted = false
let mutable isHotkeyStopped = false



let removeStartHotkey() =
    
    if isHotkeyStarted then
        let mutable ctrlfound = false
        let mutable shiftfound = false
        let mutable ctrlindex = -1
        let mutable shiftindex = -1
        for i = 0 to macroBuf.Count - 1 do
            if macroBuf.[i] :? SysKeyPressEvent then
                if (macroBuf.[i] :?> SysKeyPressEvent).Name = "lCtrl" && (macroBuf.[i] :?> SysKeyPressEvent).IsDown = false && ctrlfound = false then
                    ctrlindex <- i
                    ctrlfound <- true
                if (macroBuf.[i] :?> SysKeyPressEvent).Name = "lShift" && (macroBuf.[i] :?> SysKeyPressEvent).IsDown = false && shiftfound = false then
                    shiftindex <- i
                    shiftfound <- true
        if ctrlindex < shiftindex then
            macroBuf.RemoveAt(ctrlindex)
            macroBuf.RemoveAt(shiftindex - 1)
        else
            macroBuf.RemoveAt(shiftindex)
            macroBuf.RemoveAt(ctrlindex - 1)
            
    if isHotkeyStopped then
        let mutable ctrlfound = false
        let mutable shiftfound = false
        let mutable ctrlindex = -1
        let mutable shiftindex = -1
        for i = macroBuf.Count - 1 downto 0 do
            if macroBuf.[i] :? SysKeyPressEvent then
                if (macroBuf.[i] :?> SysKeyPressEvent).Name = "lCtrl" && (macroBuf.[i] :?> SysKeyPressEvent).IsDown = true && ctrlfound = false then
                    ctrlindex <- i
                    ctrlfound <- true
                if (macroBuf.[i] :?> SysKeyPressEvent).Name = "lShift" && (macroBuf.[i] :?> SysKeyPressEvent).IsDown = true && shiftfound = false then
                    shiftindex <- i
                    shiftfound <- true
        macroBuf.RemoveAt(shiftindex)
        macroBuf.RemoveAt(ctrlindex)
        
    isHotkeyStarted <- false
    isHotkeyStopped <- false
    
    
// Funkcja wyrejestrowania skrótu
let unregisterHotKey (handle: IntPtr) (hotkeyId: int) =
    UnregisterHotKey(handle, hotkeyId) |> ignore

let canParseToInt (input: string) : bool =
    let success, _ = Int32.TryParse(input)
    success

let refactorChar (input: string) : string =
    // Mapa znaków: Shifted -> Bez Shift
    let shiftMap = Dictionary<char, char>()
    shiftMap.Add('!', '1')
    shiftMap.Add('@', '2')
    shiftMap.Add('#', '3')
    shiftMap.Add('$', '4')
    shiftMap.Add('%', '5')
    shiftMap.Add('^', '6')
    shiftMap.Add('&', '7')
    shiftMap.Add('*', '8')
    shiftMap.Add('(', '9')
    shiftMap.Add(')', '0')
    shiftMap.Add('_', '-') // Underscore -> Dash
    shiftMap.Add('+', '=') // Plus -> Equal
    shiftMap.Add('{', '[') // Curly brace open -> Square bracket open
    shiftMap.Add('}', ']') // Curly brace close -> Square bracket close
    shiftMap.Add('|', '\\') // Pipe -> Backslash
    shiftMap.Add(':', ';') // Colon -> Semicolon
    shiftMap.Add('"', '\'') // Double quote -> Single quote
    shiftMap.Add('<', ',') // Less than -> Comma
    shiftMap.Add('>', '.') // Greater than -> Dot
    shiftMap.Add('?', '/') // Question mark -> Slash

    // Zamień znaki w ciągu wejściowym
    let mutable result = ""
    for c in input do
        if shiftMap.ContainsKey(c) then
            result <- result + string shiftMap.[c]
        else
            result <- result + string c

    result

let isStillPressed(event: SysKeyPressEvent): bool =
    let mutable result = false
    let mutable isResultFound = false
    for i = macroBuf.Count - 1 downto 0 do
        if macroBuf.[i] :? SysKeyPressEvent && isResultFound = false then
            let keyEvent = macroBuf.[i] :?> SysKeyPressEvent
            if keyEvent.Name = event.Name then
                result <- keyEvent.IsDown
                isResultFound <- true
                              
    result
    
    
let addWaitEvent(timestamp: int) =
    if timestamp >= 200 then
        let wait = WaitEvent(timestamp)
        macroBuf.Add(wait)
        //wait.Display()
        
    stopwatch <- Stopwatch.StartNew()   
    
    
let getCursorPosition () =
    let mutable cursorPos = System.Drawing.Point()
    if GetCursorPos(&cursorPos) then
        printfn "Pozycja kursora: X = %d, Y = %d" cursorPos.X cursorPos.Y
    else
        printfn "Nie udało się pobrać pozycji kursora."
        

let mouseHookProc = HookProc(fun nCode wParam lParam ->
    if nCode >= 0 then
        if wParam = IntPtr(WM_LBUTTONDOWN) then
            let mutable cursorPos = System.Drawing.Point()
            if GetCursorPos(&cursorPos) then
                let timestamp = stopwatch.ElapsedMilliseconds  // Pobierz czas w milisekundach
                let event = MousePressEvent("Left", cursorPos.X, cursorPos.Y)
                
                addWaitEvent((int)timestamp)
                macroBuf.Add(event)
                //event.Display()
        
        else if wParam = IntPtr(WM_RBUTTONDOWN) then
            let mutable cursorPos = System.Drawing.Point()
            if GetCursorPos(&cursorPos) then
                let timestamp = stopwatch.ElapsedMilliseconds  // Pobierz czas w milisekundach
                let event = MousePressEvent("Right", cursorPos.X, cursorPos.Y)
                
                addWaitEvent((int)timestamp)
                macroBuf.Add(event)
                //event.Display()
                
        else if wParam = IntPtr(WM_MBUTTONDOWN) then
            let mutable cursorPos = System.Drawing.Point()
            if GetCursorPos(&cursorPos) then
                let timestamp = stopwatch.ElapsedMilliseconds
                let event = MousePressEvent("Middle", cursorPos.X, cursorPos.Y)
                
                addWaitEvent((int)timestamp)
                macroBuf.Add(event)
                //event.Display()
                
        elif wParam = IntPtr(WM_MOUSEWHEEL) then
            let mutable cursorPos = System.Drawing.Point()
            if GetCursorPos(&cursorPos) && lParam <> IntPtr.Zero then
                let msllHookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam)
                
                let wheelDelta = int16(msllHookStruct.mouseData >>> 16) // Wyciągnij wartość delta dla scrolla
                let timestamp = stopwatch.ElapsedMilliseconds
                let event = MouseScrollEvent((int)wheelDelta)
                
                addWaitEvent((int)timestamp)
                macroBuf.Add(event)
                //event.Display()
      
    CallNextHookEx(mouseHookId, nCode, wParam, lParam)
)

let keyboardHookProc = HookProc(fun nCode wParam lParam ->
    if nCode >= 0 then
        // Sprawdzanie, czy klawisz został wciśnięty
        if wParam = IntPtr(WM_KEYDOWN) then
            let vkCode = Marshal.ReadInt32(lParam)  // Pobierz kod wirtualny klawisza
            let mutable keyName: string = null
            
            if (vkCode >= 48 && vkCode <= 57) || (vkCode >= 65 && vkCode <= 90) || noASCIkeys.ContainsKey(vkCode) then
                // Ustalamy nazwę klawisza
                keyName <- 
                    if noASCIkeys.ContainsKey(vkCode) then noASCIkeys.[vkCode]
                    else string (char vkCode)
                
                if macroBuf.Count > 0 then                
                    let prevEvent = macroBuf.[macroBuf.Count - 1]
                    if prevEvent :? KeyStroke then
                        keyStrokeBuf.Add(keyName)
                        let newValue = String.concat "" keyStrokeBuf
                        let newEvent = KeyStroke(newValue)
                        //(macroBuf.[macroBuf.Count - 1] :?> KeyStroke).Value <- newValue
                        //macroBuf.[macroBuf.Count - 1].Display()
                        macroBuf.RemoveAt(macroBuf.Count - 1)
                        macroBuf.Add(newEvent)
                        stopwatch <- Stopwatch.StartNew()
                    else
                        let timestamp = stopwatch.ElapsedMilliseconds
                        let event = KeyStroke(keyName)
                        
                        keyStrokeBuf.Clear()                       
                        keyStrokeBuf.Add(keyName)
                        addWaitEvent((int)timestamp)
                        macroBuf.Add(event)
                        //event.Display()
                else                
                    let timestamp = stopwatch.ElapsedMilliseconds
                    let event = KeyStroke(keyName)
                        
                    keyStrokeBuf.Clear()                       
                    keyStrokeBuf.Add(keyName)
                    addWaitEvent((int)timestamp)
                    macroBuf.Add(event)
                    //event.Display()
               
            else
                // Sprawdzenie, czy klucz istnieje w sysKeys
                if sysKeys.ContainsKey(vkCode) then
                    keyName <- sysKeys.[vkCode]
                    let isDown = true
                    let event = SysKeyPressEvent(keyName, vkCode, isDown)
                    
                    if isStillPressed(event) = false then
                        let timestamp = stopwatch.ElapsedMilliseconds  // Pobierz czas w milisekundach                        
                        addWaitEvent((int)timestamp)
                        macroBuf.Add(event)
                        //event.Display()
        
        // Sprawdzanie, czy klawisz został odpuszczony                      
        else if wParam = IntPtr(WM_KEYUP) then          
            let vkCode = Marshal.ReadInt32(lParam)
            
            if sysKeys.ContainsKey(vkCode) then
                let keyName = sysKeys.[vkCode]
                let timestamp = stopwatch.ElapsedMilliseconds            
                let isDown = false
                let event = SysKeyPressEvent(keyName, vkCode, isDown)
                addWaitEvent((int)timestamp)
                macroBuf.Add(event)
                //event.Display()
        
        // Sprawdzanie, czy wciśnięto/zwolniono klawisz systemowy
        else if wParam = IntPtr(WM_SYSKEYDOWN) || wParam = IntPtr(WM_SYSKEYUP) then
            let vkCode = Marshal.ReadInt32(lParam)
            
            if sysKeys.ContainsKey(vkCode) then
                let keyName = sysKeys.[vkCode]
                let timestamp = stopwatch.ElapsedMilliseconds
                let isDown = wParam = IntPtr(WM_SYSKEYDOWN) || wParam = IntPtr(WM_SYSKEYUP)
                let event = SysKeyPressEvent(keyName, vkCode, isDown)
                if isDown = false then
                    addWaitEvent((int)timestamp)
                    macroBuf.Add(event)
                    //event.Display()
                else
                    if isStillPressed(event) = false then
                        addWaitEvent((int)timestamp)
                        macroBuf.Add(event)
                        //event.Display()
                           
    CallNextHookEx(keyboardHookId, nCode, wParam, lParam)
)





    