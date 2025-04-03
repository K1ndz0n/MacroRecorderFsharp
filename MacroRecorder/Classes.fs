module Classes

open Structures
open System
open Microsoft.FSharp.Collections
open Microsoft.FSharp.Core
open System.Threading
open System.Collections.ObjectModel

let mutable isHalted = false
let sleep (milliseconds: int) =
    Thread.Sleep(milliseconds)


[<AbstractClass>]
type MacroEvent() =
    abstract member Display: unit -> string
    abstract member Playback: unit -> unit
    abstract member Clone: unit -> MacroEvent

type SysKeyPressEvent(name: string, vkCode: int, isDown: bool) =
    inherit MacroEvent()
    member val Name = name with get, set
    member val VkCode = vkCode with get, set
    member val IsDown = isDown with get, set

    override this.Display() =
        let status = if this.IsDown then "Wciśnięty" else "Puszczony"
        "Klawisz systemowy: " + this.Name + " " + status

    override this.Playback() =
        if isDown then
            keybd_event(byte this.VkCode, 0uy, 0u, IntPtr.Zero)   // Naciśnięcie
        else
            keybd_event(byte this.VkCode, 0uy, KEYEVENTF_KEYUP, IntPtr.Zero) // Zwolnienie
            
    override this.Clone() =
        SysKeyPressEvent(this.Name, this.VkCode, this.IsDown) :> MacroEvent
    
    
// Zdarzenie kliknięcia myszy
type MousePressEvent(button: string, posX: int, posY: int) =
    inherit MacroEvent()
    member val Button = button with get, set
    member val PosX = posX with get, set
    member val PosY = posY with get, set

    override this.Display() =
        if this.Button = "Left" then
            "Lewy przycisk myszy (" + (string)this.PosX + ", " + (string)this.PosY + ")"
        else if this.Button = "Right" then
            "Prawy przycisk myszy (" + (string)this.PosX + ", " + (string)this.PosY + ")"
        else
            "Środkowy przycisk myszy (" + (string)this.PosX + ", " + (string)this.PosY + ")"
            
    override this.Playback() =
        match this.Button with
        | "Left" ->
            System.Windows.Forms.Cursor.Position <- System.Drawing.Point(this.PosX, this.PosY)
            mouse_event(MOUSEEVENTF_LEFTDOWN, uint32 this.PosX, uint32 this.PosY, 0u, IntPtr.Zero)
            mouse_event(MOUSEEVENTF_LEFTUP, uint32 this.PosX, uint32 this.PosY, 0u, IntPtr.Zero)

        | "Right" ->
            System.Windows.Forms.Cursor.Position <- System.Drawing.Point(this.PosX, this.PosY)
            mouse_event(MOUSEEVENTF_RIGHTDOWN, uint32 this.PosX, uint32 this.PosY, 0u, IntPtr.Zero)
            mouse_event(MOUSEEVENTF_RIGHTUP, uint32 this.PosX, uint32 this.PosY, 0u, IntPtr.Zero)

        | "Middle" ->
            System.Windows.Forms.Cursor.Position <- System.Drawing.Point(this.PosX, this.PosY)
            mouse_event(uint32 MOUSEEVENTF_MIDDLEDOWN, uint32 this.PosX, uint32 this.PosY, 0u, IntPtr.Zero)
            mouse_event(uint32 MOUSEEVENTF_MIDDLEUP, uint32 this.PosX, uint32 this.PosY, 0u, IntPtr.Zero)
            
    override this.Clone() =
        MousePressEvent(this.Button, this.PosX, this.PosY) :> MacroEvent
  
// zdarzenie scrolla
type MouseScrollEvent(delta: int) =
    inherit MacroEvent()
    member val Delta = delta with get, set
    override this.Display() =
        if this.Delta > 0 then
            "Kółko myszy w górę"
        else
            "Kółko myszy w dół"
        
    override this.Playback() =
        mouse_event(MOUSEEVENTF_WHEEL, 0u, 0u, (uint32)this.Delta, IntPtr.Zero)
        sleep(20)
        
    override this.Clone() =
        MouseScrollEvent(this.Delta) :> MacroEvent
            
// Wiele klikniec w klawiature np wpisanie tekstu
type KeyStroke(value: string) =
    inherit MacroEvent()
    member val Value = value with get, set // Ciąg znaków

    override this.Display() =
        "Ciąg znaków: " + this.Value
        
    override this.Playback() =
        for key in this.Value do
            // Obsługa znaków ASCII (litery i cyfry)
            if ((int)key >= 48 && (int)key <= 57) || ((int)key >= 65 && (int)key <= 90) then
                let vkCode = int key
                keybd_event(byte vkCode, 0uy, 0u, IntPtr.Zero)   // Naciśnięcie
                keybd_event(byte vkCode, 0uy, KEYEVENTF_KEYUP, IntPtr.Zero) // Zwolnienie
            else
                // Obsługa znaków nie-ASCII
                let found = 
                    noASCIkeys 
                    |> Seq.tryFind (fun kvp -> kvp.Value = string key)

                match found with
                | Some kvp ->
                    let vkCode = kvp.Key
                    keybd_event(byte vkCode, 0uy, 0u, IntPtr.Zero)   // Naciśnięcie
                    keybd_event(byte vkCode, 0uy, KEYEVENTF_KEYUP, IntPtr.Zero) // Zwolnienie
                | None ->
                    printfn "Nieobsługiwany znak: %c" key
                
    override this.Clone() =
        KeyStroke(this.Value) :> MacroEvent

                                           
type WaitEvent(timeAmount: int) =
    inherit MacroEvent()
    member val TimeAmount = timeAmount with get, set
    
    override this.Display() =
        "Oczekiwanie: " + string this.TimeAmount + " ms"
        
    override this.Playback() =
        sleep(this.TimeAmount)
       
    override this.Clone() =
        WaitEvent(this.TimeAmount) :> MacroEvent
        
                
type Macro(name: string, events: ObservableCollection<MacroEvent>) =
    member val Name = name with get, set
    member val Events = events with get, set
    member val KeyBind: string = null with get, set
    
    member this.Playback() =
        for i in this.Events do
            if isHalted = false then i.Playback()
            else ()
        isHalted <- false
            
    member this.Display() =
        if this.KeyBind = null then
            this.Name
        else
            this.Name + ", skrót: lCtrl + lShift + " + this.KeyBind
        


